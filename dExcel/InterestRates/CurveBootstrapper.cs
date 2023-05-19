using dExcel.CommonEnums;
using dExcel.Dates;
using dExcel.ExcelUtils;
using dExcel.Utilities;
using ExcelDna.Integration;
using Omicron;
using QL = QuantLib;

namespace dExcel.InterestRates;

/// <summary>
/// A class containing a collection of interest rate curve bootstrapping utilities.
/// </summary>
public static class CurveBootstrapper
{

    /// <summary>
    /// Lists all available interpolation methods for interest rate curve bootstrapping.
    /// </summary>
    /// <returns>A column of all available interpolation methods for interest rate curve bootstrapping.</returns>
    [ExcelFunction(
        Name = "d.Curve_GetAvailableBootstrappingInterpolationMethods",
        Description = "Returns all available interpolation methods for interest rate curve bootstrapping.",
        Category = "∂Excel: Interest Rates")]
    public static object GetAvailableBootstrappingInterpolationMethods()
    {
        Array methods = Enum.GetValues(typeof(CurveInterpolationMethods));
        object[,] output = new object[methods.Length + 1, 1];
        output[0, 0] = "IR Bootstrapping Interpolation Methods";
        int i = 0;
        foreach (CurveInterpolationMethods method in methods)
        {
            output[i++, 0] = method.ToString();
        }
        
        return output;
    }

    /// <summary>
    /// Gets the IBOR index for the given name and tenor and can also apply the forecast curve if supplied.
    /// </summary>
    /// <param name="indexName">The index name.</param>
    /// <param name="indexTenor">The index tenor e.g., "3M", "1Y" etc.</param>
    /// <param name="forecastCurve">The forecast curve for the index.</param>
    /// <returns>The IBOR index, if successful, otherwise null.</returns>
    public static QL.IborIndex? GetIborIndex(
        string? indexName,
        string? indexTenor,
        QL.RelinkableYieldTermStructureHandle? forecastCurve = null)
    {
        if (!Enum.TryParse(indexName, out RateIndices iborName))
        {
            return null;
        }
        
        QL.IborIndex? index;
        if (forecastCurve is null)
        {
            index =
                iborName switch
                {
                    RateIndices.EURIBOR => new QL.Euribor(new QL.Period(indexTenor)),
                    RateIndices.FEDFUND => new QL.FedFunds(),
                    RateIndices.JIBAR => new QL.Jibar(new QL.Period(indexTenor)),
                    RateIndices.USD_LIBOR => new QL.USDLibor(new QL.Period(indexTenor)),
                    _ => null,
                };
        }
        else 
        {
            index =
                iborName switch
                {
                    RateIndices.EURIBOR => new QL.Euribor(new QL.Period(indexTenor), forecastCurve),
                    RateIndices.FEDFUND => new QL.FedFunds(forecastCurve),
                    RateIndices.JIBAR => new QL.Jibar(new QL.Period(indexTenor), forecastCurve),
                    RateIndices.USD_LIBOR => new QL.USDLibor(new QL.Period(indexTenor), forecastCurve),
                    _ => null,
                };
        }

        return index;
    }

    /// <summary>
    /// Bootstraps an interest rate curve. It supports multi-curve bootstrapping.
    /// Available Indices: EURIBOR, FEDFUND (OIS), JIBAR, USD-LIBOR.
    /// </summary>
    /// <param name="handle">The 'handle' or name used to refer to the object in memory.
    /// Each object in a workbook must have a unique handle.</param>
    /// <param name="curveParameters">The parameters required to construct the curve.</param>
    /// <param name="customRateIndex">(Optional)A custom rate index.</param>
    /// <param name="instrumentGroups">The list of instrument groups used in the bootstrapping.</param>
    /// <returns>A handle to a bootstrapped curve.</returns>
    [ExcelFunction(
        Name = "d.Curve_Bootstrap",
        Description = "Bootstraps a single currency interest rate curve. Supports multi-curve bootstrapping.",
        Category = "∂Excel: Interest Rates")]
    public static string Bootstrap(
        [ExcelArgument(Name = "Handle", Description = DescriptionUtils.Handle)]
        string handle,
        [ExcelArgument(
            Name = "Curve Parameters",
            Description = 
                "The curves parameters: " +
                "'BaseDate', 'RateIndexName', 'RateIndexTenor', 'Interpolation', (Optional)'DiscountCurveHandle', " +
                "(Optional)'AllowExtrapolation' (Default = False)")]
        object[,] curveParameters,
        [ExcelArgument(
            Name = "(Optional)Custom Rate Index",
            Description =
                "Only populate this if you have NOT supplied a 'RateIndexName' in the curve parameters.")]
        object[,]? customRateIndex = null,
        [ExcelArgument(
            Name = "Instrument Groups",
            Description = "The instrument groups used to bootstrap the curve e.g., 'Deposits', 'FRAs', 'Swaps'.")]
        params object[] instrumentGroups)
    {
        int columnHeaderIndex = ExcelTableUtils.GetRowIndex(curveParameters, "Parameter");
        
        DateTime baseDate =
            ExcelTableUtils.GetTableValue<DateTime>(curveParameters, "Value", "BaseDate", columnHeaderIndex);
        
        if (baseDate == default)
        {
            return CommonUtils.CurveParameterMissingErrorMessage(nameof(baseDate).ToUpper());
        }

        QL.Settings.instance().setEvaluationDate(baseDate.ToQuantLibDate());

        string? rateIndexName =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "RateIndexName", columnHeaderIndex);
        
        if (rateIndexName is null && customRateIndex is null)
        {
            return CommonUtils.CurveParameterMissingErrorMessage(nameof(rateIndexName).ToUpper());
        }

        string? rateIndexTenor =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "RateIndexTenor", columnHeaderIndex);
        
        if (rateIndexTenor is null && customRateIndex is null && rateIndexName != "FEDFUND")
        {
            return CommonUtils.CurveParameterMissingErrorMessage(nameof(rateIndexTenor).ToUpper());
        }

        string? interpolation =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "Interpolation", columnHeaderIndex);
        
        if (interpolation == null)
        {
            return CommonUtils.CurveParameterMissingErrorMessage(nameof(interpolation).ToUpper());
        }
        
        string? discountCurveHandle =
            ExcelTableUtils.GetTableValue<string?>(curveParameters, "Value", "DiscountCurveHandle", columnHeaderIndex);
        
        bool? allowExtrapolation =
            ExcelTableUtils.GetTableValue<bool?>(curveParameters, "Value", "AllowExtrapolation", columnHeaderIndex);
        
        if (allowExtrapolation == null)
        {
            allowExtrapolation = false;
        }
        
        QL.IborIndex? rateIndex = GetIborIndex(rateIndexName, rateIndexTenor, null);
        // else
        // {
        //     string? tenor = ExcelTable.GetTableValue<string>(customRateIndex, "Value", "Tenor");
        //     int? settlementDays = ExcelTable.GetTableValue<int>(customRateIndex, "Value", "SettlementDay");
        //     string? currency = ExcelTable.GetTableValue<string>(customRateIndex, "Value", "Currency");
        //     (Currency x, Calendar y) z =
        //         currency switch
        //         {
        //             // TODO: Use reflection here.
        //             "USD" => (new USDCurrency(), new UnitedStates()),
        //             "ZAR" => (new ZARCurrency(), new SouthAfrica()),
        //             _ => throw new NotImplementedException(),
        //         };
        //
        //     rateIndex = new IborIndex("Test", new Period(tenor), settlementDays, z.x, z.y, )
        // }

        if (rateIndex is null)
        {
            return CommonUtils.DExcelErrorMessage($"Unsupported rate index: {rateIndexName}");
        }

        QL.RateHelperVector rateHelpers = new();

        foreach (object instrumentGroup in instrumentGroups)
        {
            object[,] instruments = (object[,]) instrumentGroup;
            string? instrumentType = ExcelTableUtils.GetTableLabel(instruments);
            List<string>? tenors = ExcelTableUtils.GetColumn<string>(instruments, "Tenors");
            List<string>? fraTenors = ExcelTableUtils.GetColumn<string>(instruments, "FraTenors");
            List<DateTime>? endDates = ExcelTableUtils.GetColumn<DateTime>(instruments, "EndDates");
            List<double>? rates = ExcelTableUtils.GetColumn<double>(instruments, "Rates");
            List<bool>? includeInstruments = ExcelTableUtils.GetColumn<bool>(instruments, "Include");

            if (includeInstruments is null)
            {
                continue;
            }

            int instrumentCount = includeInstruments.Count;

            if (string.Equals(instrumentType, "Deposits", StringComparison.OrdinalIgnoreCase))
            {
                for (int i = 0; i < instrumentCount; i++)
                {
                    if (rates is null)
                    {
                        return CommonUtils.DExcelErrorMessage("Deposit rates missing.");
                    }

                    if (tenors is null)
                    {
                        return CommonUtils.DExcelErrorMessage("Deposit tenors missing");
                    }

                    if (includeInstruments[i])
                    {
                        rateHelpers.Add(
                            new QL.DepositRateHelper(
                                rate: rates[i],
                                tenor: new QL.Period(tenors[i]),
                                fixingDays: rateIndex.fixingDays(),
                                calendar: rateIndex.fixingCalendar(),
                                convention: rateIndex.businessDayConvention(),
                                endOfMonth: rateIndex.endOfMonth(),
                                dayCounter: rateIndex.dayCounter()));
                    }
                }
            }
            else if (instrumentType.IgnoreCaseEquals("FRAs"))
            {
                for (int i = 0; i < instrumentCount; i++)
                {
                    if (includeInstruments[i])
                    {
                        if (fraTenors is null)
                        {
                            return CommonUtils.DExcelErrorMessage("FRA tenors missing.");
                        }

                        if (rates is null)
                        {
                            return CommonUtils.DExcelErrorMessage("FRA rates missing.");
                        }

                        rateHelpers.Add(
                            new QL.FraRateHelper(
                                rate: new QL.QuoteHandle(new QL.SimpleQuote(rates[i])),
                                periodToStart: new QL.Period(fraTenors[i]),
                                iborIndex: rateIndex));
                    }
                }
            }
            else if (instrumentType.IgnoreCaseEquals("Interest Rate Swaps"))
            {
                for (int i = 0; i < instrumentCount; i++)
                {
                    if (rates is null)
                    {
                        return CommonUtils.DExcelErrorMessage("Swap rates missing.");
                    }

                    if (tenors is null)
                    {
                        return CommonUtils.DExcelErrorMessage("Swap tenors missing");
                    }

                    if (includeInstruments[i])
                    {
                        QL.RelinkableYieldTermStructureHandle discountCurve = new();
                        if (discountCurveHandle != null)
                        {
                            QL.YieldTermStructure? yieldTermStructure = CurveUtils.GetCurveObject(discountCurveHandle);
                            discountCurve.linkTo(yieldTermStructure);
                        }
                       
                        QL.QuoteHandle quoteHandle = new(new QL.SimpleQuote(rates[i]));
                        rateHelpers.Add(
                            new QL.SwapRateHelper(
                                rate: quoteHandle,
                                tenor: new QL.Period(tenors[i]),
                                calendar: rateIndex.fixingCalendar(),
                                fixedFrequency: rateIndex.tenor().frequency(),
                                fixedConvention: rateIndex.businessDayConvention(),
                                fixedDayCount: rateIndex.dayCounter(),
                                index: rateIndex,
                                spread: new QL.QuoteHandle(new QL.SimpleQuote(0)),
                                fwdStart: new QL.Period(0, QL.TimeUnit.Months),
                                discountingCurve: discountCurve));
                    }
                }
            }
            else if (instrumentType.Equals("OISs", StringComparison.OrdinalIgnoreCase))
            {
                if (rates is null)
                {
                    return CommonUtils.DExcelErrorMessage("OIS rates missing.");
                }
                
                for (int i = 0; i < instrumentCount; i++)
                {
                    if (includeInstruments[i])
                    {
                        rateHelpers.Add(
                            new QL.OISRateHelper(
                                settlementDays: 2, 
                                tenor: new QL.Period(tenors?[i]),
                                rate: new QL.QuoteHandle(new QL.SimpleQuote(rates[i])), 
                                index: rateIndex as QL.OvernightIndex));
                    }
                }
            }
        }

        QL.YieldTermStructure? termStructure =
            BootstrapCurveFromRateHelpers(
                rateHelpers: rateHelpers, 
                referenceDate: baseDate, 
                dayCountConvention: rateIndex.dayCounter(), 
                interpolation: interpolation);
        
        if (termStructure is null)
        {
            return CommonUtils.DExcelErrorMessage($"Unknown interpolation method: '{interpolation}'");
        }
        
        if ((bool)allowExtrapolation)
        {
            termStructure.enableExtrapolation();
        }
        
        CurveDetails curveDetails = 
            new(termStructure: termStructure, 
                dayCountConvention: rateIndex.dayCounter(), 
                interpolation: interpolation, 
                discountFactorDates: null,  
                discountFactors: null, 
                instrumentGroups: instrumentGroups);
        
        DataObjectController dataObjectController = DataObjectController.Instance;
        return dataObjectController.Add(handle, curveDetails);
    }
    
    /// <summary>
    /// Bootstraps a curve (single or multi-curve) given a list of rate helpers.
    /// </summary>
    /// <param name="rateHelpers">Rate helpers.</param>
    /// <param name="referenceDate">The curve base date.</param>
    /// <param name="dayCountConvention">The day count convention.</param>
    /// <param name="interpolation">The interpolation method.</param>
    /// <returns>A bootstrapped term structure, if it succeeds, otherwise null.</returns>
    public static QL.YieldTermStructure? BootstrapCurveFromRateHelpers(
        QL.RateHelperVector rateHelpers,
        DateTime referenceDate,
        QL.DayCounter dayCountConvention,
        string interpolation)
    {
        QL.Date curveDate = referenceDate.ToQuantLibDate();
        if (interpolation.IgnoreCaseEquals(CurveInterpolationMethods.Flat_On_ForwardRates))
        {
            return new QL.PiecewiseFlatForward(curveDate, rateHelpers, dayCountConvention);
        }

        if (interpolation.IgnoreCaseEquals(CurveInterpolationMethods.CubicSpline_On_DiscountFactors))
        {
            return new QL.PiecewiseSplineCubicDiscount(curveDate, rateHelpers, dayCountConvention);
        }

        if (interpolation.IgnoreCaseEquals(CurveInterpolationMethods.Exponential_On_DiscountFactors))
        {
            return new QL.PiecewiseLogLinearDiscount(curveDate, rateHelpers, dayCountConvention);
        }
        
        if (interpolation.IgnoreCaseEquals(CurveInterpolationMethods.LogCubic_On_DiscountFactors))
        {
            return new QL.PiecewiseLogCubicDiscount(curveDate, rateHelpers, dayCountConvention);
        }

        if (interpolation.IgnoreCaseEquals(CurveInterpolationMethods.NaturalLogCubic_On_DiscountFactors))
        {
            return new QL.PiecewiseNaturalLogCubicDiscount(curveDate, rateHelpers, dayCountConvention);
        }

        if (interpolation.IgnoreCaseEquals(CurveInterpolationMethods.Cubic_On_ZeroRates))
        {
            return new QL.PiecewiseCubicZero(curveDate, rateHelpers, dayCountConvention);
        }
        
        if (interpolation.IgnoreCaseEquals(CurveInterpolationMethods.Linear_On_ZeroRates))
        {
            return new QL.PiecewiseLinearZero(curveDate, rateHelpers, dayCountConvention);
        }
        
        if (interpolation.IgnoreCaseEquals(CurveInterpolationMethods.NaturalCubic_On_ZeroRates))
        {
            return new QL.PiecewiseNaturalCubicZero(curveDate, rateHelpers, dayCountConvention);
        }
        
        return null;
    }
    
    /// <summary>
    /// Extracts and bootstraps a curve from the Omicron database.
    /// </summary>
    /// <param name="handle">The 'handle' or name used to refer to the object in memory.
    /// Each object in a workbook must have a a unique handle.</param>
    /// <param name="curveName">The name of the curve in Omicron. Current available options are:
    /// 'ZAR-Swap', 'USD-OIS'</param>
    /// <param name="baseDate"></param>
    /// <param name="interpolation"></param>
    /// <returns></returns>
    [ExcelFunction(
        Name = "d.Curve_Get",
        Description = "Extracts and bootstraps a curve from the Omicron database.",
        Category = "∂Excel: Interest Rates")]
    public static string Get(
        [ExcelArgument(Name = "Handle", Description = DescriptionUtils.Handle)]
        string handle,
        [ExcelArgument(
            Name = "Curve Name",
            Description = 
                "The name of the curve in Omicron. Current available options are:\n" +
                "• USD-OIS\n" +
                "• ZAR-Swap")]
        string curveName,
        [ExcelArgument(
            Name = "Base Date",
            Description = "The base date of the curve i.e., the date for which to extract the curve.")]
        DateTime baseDate,
        [ExcelArgument(
            Name = "DF Interpolation",
            Description = 
                "(Optional)The discount factor interpolation style.\n" +
                "Default = 'Exponential_On_DiscountFactors'.")]
        string interpolation = "")
    {
        interpolation = 
            interpolation != "" ? interpolation : CurveInterpolationMethods.Exponential_On_DiscountFactors.ToString();
        
        // TODO: List the available types of interpolation.
        // One could create more complicated abstract code for mapping from quotes to 2d tables but I would advise
        // against this.
        string rateIndexName = "";
        string rateIndexTenor = "";
        switch (curveName.ToUpper())
        {
            case "ZAR-SWAP":
                rateIndexName = "JIBAR"; 
                rateIndexTenor = "3m";
                break;
            case "USD-OIS":
                rateIndexName = "FEDFUND";
                rateIndexTenor = "1d";
                break;
        }

        List<QuoteValue> quoteValues;
        try
        {
            quoteValues =
                OmicronUtils.OmicronUtils.GetSwapCurveQuotes(rateIndexName, null, 1, baseDate.ToString("yyyy-MM-dd"));
        }
        catch (Exception ex)
        {
            if (!NetworkUtils.GetVpnConnectionStatus())
            {
                return CommonUtils.DExcelErrorMessage("Not connected to Deloitte network/VPN.");
            }

            return CommonUtils.DExcelErrorMessage($"Unknown error. {ex.Message}");
        }

        object[,] curveParameters =
        {
            {"CurveUtils Parameters", ""},
            {"Parameter", "Value"},
            {"BaseDate", baseDate.ToOADate()},
            {"RateIndexName", rateIndexName},
            {"RateIndexTenor", rateIndexTenor},
            {"Interpolation", interpolation},
        };

        List<QuoteValue> deposits = quoteValues.Where(x => x.Type.GetType() == typeof(RateIndex)).ToList();
        deposits = deposits.OrderBy(x => ((RateIndex)x.Type).Tenor, new TenorComparer()).ToList();
        object[,] depositInstruments = new object[deposits.Count + 2, 4];
        depositInstruments[0, 0] = "Deposits";
        depositInstruments[1, 0] = "Tenors";
        depositInstruments[1, 1] = "RateIndex";
        depositInstruments[1, 2] = "Rates";
        depositInstruments[1, 3] = "Include";

        int row = 2;
        foreach (QuoteValue deposit in deposits)
        {
            depositInstruments[row, 0] = ((RateIndex) deposit.Type).Tenor.ToString();
            depositInstruments[row, 1] = ((RateIndex) deposit.Type).Name;
            depositInstruments[row, 2] = deposit.Value;
            depositInstruments[row, 3] = "TRUE";
            row++;
        }

        List<QuoteValue> fras = quoteValues.Where(x => x.Type.GetType() == typeof(Fra)).ToList();
        fras = fras.OrderBy(x => ((Fra)x.Type).Tenor, new TenorComparer()).ToList();
        object[,] fraInstruments = new object[fras.Count + 2, 4];
        row = 2;
        fraInstruments[0, 0] = "FRAs";
        fraInstruments[1, 0] = "FraTenors";
        fraInstruments[1, 1] = "RateIndex";
        fraInstruments[1, 2] = "Rates";
        fraInstruments[1, 3] = "Include";

        foreach (QuoteValue fra in fras)
        {
            // TODO: Ensure the amount is always in months.
            fraInstruments[row, 0] = $"{((Fra) fra.Type).Tenor.Amount}x{((Fra) fra.Type).Tenor.Amount + 3}";
            fraInstruments[row, 1] = ((Fra) fra.Type).ReferenceIndex.Name;
            fraInstruments[row, 2] = fra.Value;
            fraInstruments[row, 3] = "TRUE";
            row++;
        }

        List<QuoteValue> swaps = quoteValues.Where(x => x.Type.GetType() == typeof(InterestRateSwap)).ToList();
        swaps = swaps.OrderBy(x => ((InterestRateSwap)x.Type).Tenor, new TenorComparer()).ToList();
        object[,] swapInstruments = new object[swaps.Count + 2, 4];
        swapInstruments[0, 0] = "Interest Rate Swaps";
        swapInstruments[1, 0] = "Tenors";
        swapInstruments[1, 1] = "RateIndex";
        swapInstruments[1, 2] = "Rates";
        swapInstruments[1, 3] = "Include";

        row = 2;
        foreach (QuoteValue swap in swaps)
        {
            swapInstruments[row, 0] = ((InterestRateSwap) swap.Type).Tenor.ToString();
            swapInstruments[row, 1] = ((InterestRateSwap) swap.Type).ReferenceIndex.Name;
            swapInstruments[row, 2] = swap.Value;
            swapInstruments[row, 3] = "TRUE";
            row++;
        }
        
        List<QuoteValue> overnightIndexSwaps = quoteValues.Where(x => x.Type.GetType() == typeof(Ois)).ToList();
        overnightIndexSwaps = overnightIndexSwaps.OrderBy(x => ((Ois)x.Type).Tenor, new TenorComparer()).ToList();
        object[,] oisInstruments = new object[overnightIndexSwaps.Count + 2, 3];
        oisInstruments[0, 0] = "OISs";
        oisInstruments[1, 0] = "Tenors";
        oisInstruments[1, 1] = "Rates";
        oisInstruments[1, 2] = "Include";

        row = 2;
        foreach (QuoteValue ois in overnightIndexSwaps)
        {
            oisInstruments[row, 0] = ((Ois)ois.Type).Tenor.ToString();
            oisInstruments[row, 1] = ois.Value;
            oisInstruments[row, 2] = "TRUE";
            row++;
        }

        List<object> instruments = new();
        if (deposits.Any()) instruments.Add(depositInstruments);
        if (fras.Any()) instruments.Add(fraInstruments);
        if (swaps.Any()) instruments.Add(swapInstruments);
        if (overnightIndexSwaps.Any()) instruments.Add(oisInstruments);

        return Bootstrap(handle, curveParameters, null, instruments.ToArray());
    }
}
