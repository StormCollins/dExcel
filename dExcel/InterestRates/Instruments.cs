using QL = QuantLib;
using dExcel.Dates;
using dExcel.Utilities;
using QuantLib;
using ExcelDna.Integration;
//using QuantLib;
using System.Windows.Documents;
using System;
using System.Reflection;
using MathNet.Numerics.Differentiation;
using System.Reflection.Metadata;
using Microsoft.OData.Client;
using Microsoft.Office.Interop.Excel;
using dExcel.ExcelUtils;
using FuzzySharp.Utils;
using System.Collections.Immutable;
using Omicron.Data.Model.QuoteTypes;
using Microsoft.VisualBasic.Logging;
using Omicron;
using Omicron.Data.Model;

namespace dExcel.InterestRates;

/// <summary>
/// Common interest rate instruments.
/// </summary>
public static class Instruments
{


    /// <summary>
    /// Creates a rate index object. Allows for historical fixings to be assigned, which removes the need to assign these on an instrument level.
    /// </summary>
    /// <param name="handle">The 'handle' or name used to refer to the object in memory.
    /// Each object in a workbook must have a unique handle.</param>
    /// <param name="referenceIndexName">Name of the reference floating rate index.</param>
    /// <param name="referenceIndexTenor">Floating rate index tenor.</param>
    /// <param name="curveHandle">Curve used for forecasting the floating rate index.</param>
    /// <param name="fixingDatesRange">The dates for any applicable historical fixings.</param>
    /// <param name="fixingsRange">The fixings for the corresponding fixing dates.</param>
    /// <returns>A handle to a bootstrapped curve.</returns>
    [ExcelFunction(
            Name = "d.IR_CreateReferenceRateIndex",
            Description = "Creates a rate index object. Allows for historicak fixings to be assigned.",
            Category = "∂Excel: Interest Rates")]
    public static string CreateReferenceRateIndex(
            [ExcelArgument(Name = "Handle", Description = DescriptionUtils.Handle)]
            string handle,
            [ExcelArgument(Name = "Reference Index", Description = "Name of the reference floating rate index")]
            string referenceIndexName,
            [ExcelArgument(Name = "Tenor", Description = "Floating rate index tenor.")]
            string referenceIndexTenor,
            [ExcelArgument(Name = "CurveHandle", Description = "Curve used for forecasting the floating rate index.")]
            string curveHandle,
            [ExcelArgument(Name = "(Optional) Fixing Dates", Description = "The dates for any applicable historical fixings.")]
            object fixingDatesRange = null,
            [ExcelArgument(Name = "(Optional) Fixings", Description = "The fixings for the corresponding fixing dates.")]
            object fixingsRange=null)
    {

        QL.YieldTermStructure? interestRateCurve = CurveUtils.GetCurveObject(curveHandle);

        QL.YieldTermStructureHandle yieldTermStructureHandle = new(interestRateCurve);

        QL.IborIndex? rateIndex = null;
        if (referenceIndexName is not null)
        {
            rateIndex =
            referenceIndexName.ToUpper() switch //added toupper()
            {
                "EURIBOR" => new QL.Euribor(new QL.Period(referenceIndexTenor), yieldTermStructureHandle),
                "FEDFUND" => new QL.FedFunds(yieldTermStructureHandle),
                "JIBAR" => new QL.Jibar(new QL.Period(referenceIndexTenor), yieldTermStructureHandle),
                "USD-LIBOR" => new QL.USDLibor(new QL.Period(referenceIndexTenor), yieldTermStructureHandle),
                _ => null,
            };
        }
        
        rateIndex.clearFixings();
        if (fixingDatesRange is not null && fixingDatesRange.GetType() != typeof(double))
        {
            List<QL.Date> fixingDates = new();
            List<double> fixings = new();
            for (int i = 0; i < ((object[,])fixingDatesRange).GetLength(0); i++)
            {
                fixingDates.Add(DateTime.FromOADate((double)((object[,])fixingDatesRange)[i, 0]).ToQuantLibDate());
                fixings.Add((double)((object[,])fixingsRange)[i, 0]);
            }
            rateIndex.addFixings(new QL.DateVector(fixingDates), new QL.DoubleVector(fixings));
        }

        if (fixingDatesRange is not null && fixingDatesRange.GetType() == typeof(double))
        {
            List<QL.Date> fixingDates = new();
            List<double> fixings = new();

            fixingDates.Add(DateTime.FromOADate((double)fixingDatesRange).ToQuantLibDate());
            fixings.Add((double)fixingsRange);

            rateIndex.addFixings(new QL.DateVector(fixingDates), new QL.DoubleVector(fixings));
        }

        var dataObjectController = DataObjectController.Instance;
        return dataObjectController.Add(handle, rateIndex);

    }

    /// <summary>
    /// Used to create a fixed-for-floating interest rate swap object in Excel.
    /// </summary>
    /// <param name="handle"></param>
    /// <param name="valuationDate">The valuation date.</param>
    /// <param name="effectiveDate">Start date.</param>
    /// <param name="terminationDate">Unadjusted maturity date.</param>
    /// <param name="tenor">Payment/receive frequency (assumed to be the same).</param>
    /// <param name="calendarsToParse">The calendar for the option e.g., 'South Africa' or 'ZAR'.</param>
    /// <param name="businessDayConventionToParse">'Business day convention e.g., 'FOL', 'MODFOL', 'PREC' etc.'.</param>
    /// <param name="ruleToParse">The date generation rule.</param>
    /// <param name="swapTypeToParse">'Payer' or 'Receiver'.</param>
    /// <param name="nominal">Nominal of the swap.</param>
    /// <param name="fixedRate">Rate of fixed leg.</param>
    /// <param name="dayCountConventionToParse">Day count convention e.g., 'Act360' or 'Act365'.</param>
    /// <param name="referenceRateIndexHandle">Reference Floating Rate Index.</param>
    /// <param name="spread">Spread above floating rate.</param>
    /// <param name="curveHandle">Handle of curve that will be used for discounting.</param>
    /// <returns></returns>
    [ExcelFunction(
        Name = "d.IR_CreateVanillaInterestRateSwap",
        Description = "Creates a fixed-for-floating interest rate swap object in Excel",
        Category = "∂Excel: Interest Rates")]

    /*made stating string*/
    public static string CreateVanillaInterestRateSwap(
    [ExcelArgument(Name = "Handle", Description = DescriptionUtils.Handle)]
    string handle,
    [ExcelArgument(Name = "Valuation Date", Description = "The valuation date.")]
    DateTime valuationDate,
    [ExcelArgument(Name = "Effective Date", Description = "Start date.")]
    DateTime effectiveDate,
    [ExcelArgument(Name = "Termination Date", Description = "Unadjusted maturity date.")]
    DateTime terminationDate,
    [ExcelArgument(Name = "Tenor", Description = "Frequency of cash flows")]
    string tenor,
    [ExcelArgument(Name = "Calendar", Description = "The calendar for the option e.g., 'South Africa' or 'ZAR'.")]
    string calendarsToParse,
    [ExcelArgument(Name = "Business Day Convention", Description = "Business day convention e.g., 'FOL', 'MODFOL', 'PREC' etc.")]
    string businessDayConventionToParse,
    [ExcelArgument(
        Name = "Date generation rule",
        Description = "The date generation rule. " +
                        "\n'Backward' = Start from end date and move backwards. " +
                        "\n'Forward' = Start from start date and move forwards. " +
                        "\n'IMM' = IMM dates.")]
    string ruleToParse,
    [ExcelArgument(Name = "Swap Type", Description = "'Payer' or 'Receiver'. ")]
    string swapTypeToParse,
    [ExcelArgument(Name = "Nominal", Description = "Nominal of the swap.")]
    double nominal,
    [ExcelArgument(Name = "Fixed Rate", Description = "Rate of fixed leg.")]
    double fixedRate,
    [ExcelArgument(Name = "Day Count Convention", Description = "Day count convention e.g., 'Act360' or 'Act365'.")]
    string dayCountConventionToParse,
    [ExcelArgument(Name = "Reference Rate Index Handle", Description = "Handle of the Floating Rate Index.")]
    string referenceRateIndexHandle,
    [ExcelArgument(Name = "Spread", Description = "Spread above floating rate.")]
    double spread,
    [ExcelArgument(Name = "CurveHandle", Description = "Handle of curve that will be used for discounting.")]
    string curveHandle)

    {

        Settings.instance().setEvaluationDate(valuationDate.ToQuantLibDate());

        QL.YieldTermStructure? interestRateCurve = CurveUtils.GetCurveObject(curveHandle);

        var dataObjectControllerReferenceRateIndex = DataObjectController.Instance;
        object dataObjectIndex = dataObjectControllerReferenceRateIndex.GetDataObject(referenceRateIndexHandle);
        QL.IborIndex rateIndex = (QL.IborIndex)dataObjectIndex;

        (QL.Calendar? calendar, string calendarErrorMessage) = DateUtils.ParseCalendars(calendarsToParse);
        (QL.BusinessDayConvention? businessDayConvention, string errorMessage) = DateUtils.ParseBusinessDayConvention(businessDayConventionToParse);
        QL.DayCounter? dayCountConvention = DateUtils.ParseDayCountConvention(dayCountConventionToParse);
    
        QL.DateGeneration.Rule rule = ruleToParse.ToUpper() switch
        {
            "BACKWARD" => QL.DateGeneration.Rule.Backward,
            "FORWARD" => QL.DateGeneration.Rule.Forward,
            "IMM" => QL.DateGeneration.Rule.TwentiethIMM,
            _ => QL.DateGeneration.Rule.Forward,
        };

        QL.Swap.Type type = swapTypeToParse.ToUpper() switch
        {
            "PAYER" => QL.Swap.Type.Payer,
            "RECEIVER" => QL.Swap.Type.Receiver,
        };

        QL.Date effDateConverted = effectiveDate.ToQuantLibDate();
        QL.Date termDateConverted = terminationDate.ToQuantLibDate();

        QL.Schedule schedule = new(effDateConverted, termDateConverted, new QL.Period(tenor), calendar, (QL.BusinessDayConvention)businessDayConvention, (QL.BusinessDayConvention)businessDayConvention, rule, false);

        QL.VanillaSwap vanillaSwap = new(type, nominal, schedule, fixedRate, dayCountConvention, schedule, rateIndex, spread, dayCountConvention);
        
        QL.DiscountingSwapEngine discountingSwapEngine = new(new QL.YieldTermStructureHandle(interestRateCurve));
        vanillaSwap.setPricingEngine(discountingSwapEngine);

        var dataObjectController = DataObjectController.Instance;
        return dataObjectController.Add(handle, vanillaSwap);

    }

    
    [ExcelFunction(
        Name = "d.IR_VanillaInterestRateSwap_GetPrice",
        Description = "Get pricing elements of a Vanilla Interest Rate Swap.",
        Category = "∂Excel: Interest Rates")]
    public static object VanillaInterestRateSwap_GetPrice(string handle, bool cashFlows = false)
    {
        var dataObjectController = DataObjectController.Instance;
        object dataObject = dataObjectController.GetDataObject(handle);
        object output;

        
        if (dataObject.GetType() == typeof(QL.VanillaSwap))
        {
            QL.VanillaSwap vanillaSwap = (QL.VanillaSwap)dataObject;
            if (cashFlows)
            {
                object[,] outputTemp = new object[vanillaSwap.fixedLeg().Count, 20];
                QL.Date tempFixingDate;
                QL.Date tempFixingStartDate;
                for (int i = 0; i < vanillaSwap.fixedLeg().Count; i++)
                {
                    
                  
                    //first set of days relate to accrual period. this should be same for fixed and float legs, used fix leg (no specific reason)  to access the data
                    outputTemp[i, 0] = QL.NQuantLibc.as_fixed_rate_coupon(vanillaSwap.fixedLeg()[i]).accrualStartDate().ToDateTime().ToOADate(); 
                    outputTemp[i, 1] = QL.NQuantLibc.as_fixed_rate_coupon(vanillaSwap.fixedLeg()[i]).accrualEndDate().ToDateTime().ToOADate();
                    outputTemp[i, 2] = QL.NQuantLibc.as_fixed_rate_coupon(vanillaSwap.fixedLeg()[i]).accrualDays();

                    //forecast period can be different from accrual, access this through floating leg

                    tempFixingDate = QL.NQuantLibc.as_floating_rate_coupon(vanillaSwap.floatingLeg()[i]).fixingDate(); 
                    outputTemp[i, 3] = tempFixingDate.ToDateTime().ToOADate();                                                                   //actual fixing date
                    tempFixingStartDate = QL.NQuantLibc.as_floating_rate_coupon(vanillaSwap.floatingLeg()[0]).index().valueDate(tempFixingDate);
                    outputTemp[i, 4] = tempFixingStartDate.ToDateTime().ToOADate();                                                             //traditional forecast start date, i.e. start of underlying NCD                    
                    outputTemp[i, 5] = QL.NQuantLibc.as_floating_rate_coupon(vanillaSwap.floatingLeg()[0]).index().maturityDate(tempFixingStartDate).ToDateTime().ToOADate(); //traditional forecast end date, i.e. maturity date of underlying NCD

                    outputTemp[i, 6] = (QL.NQuantLibc.as_floating_rate_coupon(vanillaSwap.floatingLeg()[i]).index().fixingCalendar()).adjust(QL.NQuantLibc.as_fixed_rate_coupon(vanillaSwap.fixedLeg()[i]).accrualStartDate()).ToOaDate(); //actual forecast start
                    outputTemp[i, 7] = (QL.NQuantLibc.as_floating_rate_coupon(vanillaSwap.floatingLeg()[i]).index().fixingCalendar()).adjust(QL.NQuantLibc.as_fixed_rate_coupon(vanillaSwap.fixedLeg()[i]).accrualEndDate()).ToOaDate(); //actual forecast end

                    //payment date. this should be same for fixed and float legs, used fix leg (no specific reason)  to access the data
                    outputTemp[i, 8] = vanillaSwap.fixedLeg()[i].date().ToDateTime().ToOADate();

                    outputTemp[i, 9] = vanillaSwap.nominal();
                    outputTemp[i, 10] = QL.NQuantLibc.as_fixed_rate_coupon(vanillaSwap.fixedLeg()[i]).rate(); //fixed rate
                    outputTemp[i, 11] = vanillaSwap.fixedLeg()[i].amount(); //fixed projected cf

                    outputTemp[i, 12] = vanillaSwap.nominal();
                    outputTemp[i, 13] = QL.NQuantLibc.as_floating_rate_coupon(vanillaSwap.floatingLeg()[i]).index().fixing(tempFixingDate); //floating rate over NCD period
                    outputTemp[i, 14] = QL.NQuantLibc.as_floating_rate_coupon(vanillaSwap.floatingLeg()[i]).adjustedFixing(); //adjusting floating rate
                    outputTemp[i, 15] = QL.NQuantLibc.as_floating_rate_coupon(vanillaSwap.floatingLeg()[i]).spread(); //spread over floating rate
                    outputTemp[i, 16] = QL.NQuantLibc.as_floating_rate_coupon(vanillaSwap.floatingLeg()[i]).gearing(); //spread over floating rate
                    outputTemp[i, 17] = QL.NQuantLibc.as_floating_rate_coupon(vanillaSwap.floatingLeg()[i]).convexityAdjustment(); //spread over floating rate
                    outputTemp[i, 18] = QL.NQuantLibc.as_floating_rate_coupon(vanillaSwap.floatingLeg()[i]).rate(); //total rate
                    outputTemp[i, 19] = QL.NQuantLibc.as_floating_rate_coupon(vanillaSwap.floatingLeg()[i]).amount(); //float projected CF


                }
                output = outputTemp;
            }

            else
            {
                output = vanillaSwap.NPV();
            }
        }

        else
        {
            return CommonUtils.DExcelErrorMessage("Unknown interest rate swap type.");
        }

        return output;
    }


    /// <summary>
    /// Used to create a non-standard fixed-for-floating interest rate swap object in Excel.
    /// </summary>
    /// <param name="handle"></param>
    /// <param name="valuationDate">The valuation date.</param>
    /// <param name="swapTypeToParse">'Payer' or 'Receiver'.</param>
    /// <param name="fixedRate">Rate of fixed leg.</param>
    /// <param name="dayCountConventionToParse">Day count convention e.g., 'Act360' or 'Act365'.</param>
    /// <param name="referenceRateIndexHandle">Reference Floating Rate Index.</param>
    /// <param name="spread">Spread above floating rate.</param>
    /// <param name="curveHandle">Handle of curve that will be used for discounting.</param>
    /// <param name="cashFlowDatesRange">Start dates of each of the interest accrual periods.</param>
    /// <param name="nominalRangeFloatLeg">Vector containing nominals of float leg.</param>
    /// <param name="nominalRangeFixedLeg">Vector containing nominals of fixed leg.</param>
    /// <returns></returns>
    [ExcelFunction(
        Name = "d.IR_CreateNonStandardInterestRateSwap",
        Description = "Creates a non-standard fixed-for-floating interest rate swap object in Excel",
        Category = "∂Excel: Interest Rates")]

    /*made stating string*/
    public static string CreateNonStandardInterestRateSwap(
    [ExcelArgument(Name = "Handle", Description = DescriptionUtils.Handle)]
    string handle,
    [ExcelArgument(Name = "Valuation Date", Description = "The valuation date.")]
    DateTime valuationDate,
    [ExcelArgument(Name = "Swap Type", Description = "'Payer' or 'Receiver'. ")]
    string swapTypeToParse,
    [ExcelArgument(Name = "Fixed Rate", Description = "Rate of fixed leg.")]
    double fixedRate,
    [ExcelArgument(Name = "Day Count Convention", Description = "Day count convention e.g., 'Act360' or 'Act365'.")]
    string dayCountConventionToParse,
    [ExcelArgument(Name = "Reference Rate Index Handle", Description = "Handle of the Floating Rate Index.")]
    string referenceRateIndexHandle,
    [ExcelArgument(Name = "Spread", Description = "Spread above floating rate.")]
    double spread,
    [ExcelArgument(Name = "CurveHandle", Description = "Handle of curve that will be used for forecasting and discounting.")]
    string curveHandle,
    [ExcelArgument(Name = "Cash flow dates", Description = "Dates used to infer interest accrual periods. Should be one more than nominals.")]
    object cashFlowDatesRange,
    [ExcelArgument(Name = "NominalsFloat", Description = "Vector containing nominals of float leg.")]
    object nominalRangeFloatLeg,
    [ExcelArgument(Name = "NominalsFixed", Description = "Vector containing nominals of fixed leg.")]
    object nominalRangeFixedLeg)

    {

        Settings.instance().setEvaluationDate(valuationDate.ToQuantLibDate());

        QL.YieldTermStructure? interestRateCurve = CurveUtils.GetCurveObject(curveHandle);

        var dataObjectControllerReferenceRateIndex = DataObjectController.Instance;
        object dataObjectIndex = dataObjectControllerReferenceRateIndex.GetDataObject(referenceRateIndexHandle);
        QL.IborIndex rateIndex = (QL.IborIndex)dataObjectIndex;

        QL.DayCounter? dayCountConvention = DateUtils.ParseDayCountConvention(dayCountConventionToParse);

        QL.Swap.Type type = swapTypeToParse.ToUpper() switch
        {
            "PAYER" => QL.Swap.Type.Payer,
            "RECEIVER" => QL.Swap.Type.Receiver,
        };

        QL.YieldTermStructureHandle yieldTermStructureHandle = new(interestRateCurve);

        QL.DiscountingSwapEngine discountingSwapEngine = new(new QL.YieldTermStructureHandle(interestRateCurve));

        //convert excel bespoke cash flow inputs into lists

        List<QL.Date> cashFlowDates = new();
        List<double> nominalFloatLeg = new();
        List<double> nominalFixedLeg = new();
        List<double> spreadVector = new();
        List<double> gearingVector = new();
        List<double> fixedRateVector = new();

        //cash flow dates always needs to contain at least two elements, so we firstly loop through this
        for (int i = 0; i < ((object[,])cashFlowDatesRange).GetLength(0); i++)
        {
            cashFlowDates.Add(DateTime.FromOADate((double)((object[,])cashFlowDatesRange)[i, 0]).ToQuantLibDate());
        }

        //we then loop through the rest of the elements which should be one element per cash flow date
            if (nominalRangeFloatLeg.GetType() != typeof(double))
        {

            for (int i = 0; i < ((object[,])nominalRangeFloatLeg).GetLength(0); i++)
            {
                nominalFloatLeg.Add((double)((object[,])nominalRangeFloatLeg)[i, 0]);
                nominalFixedLeg.Add((double)((object[,])nominalRangeFixedLeg)[i, 0]);
                spreadVector.Add(spread);
                gearingVector.Add(1); //not using gearing feature
                fixedRateVector.Add(fixedRate);
            }

        }

        if (nominalRangeFloatLeg.GetType() == typeof(double))
        {
            nominalFloatLeg.Add((double)nominalRangeFloatLeg);
            nominalFixedLeg.Add((double)nominalRangeFixedLeg);
            spreadVector.Add(spread);
            gearingVector.Add(0);
            fixedRateVector.Add(fixedRate);

        }

        QL.NonstandardSwap nonstandardSwap = new(type, new QL.DoubleVector(nominalFixedLeg), new QL.DoubleVector(nominalFloatLeg), new QL.Schedule(new QL.DateVector(cashFlowDates)), new QL.DoubleVector(fixedRateVector), dayCountConvention, new QL.Schedule(new QL.DateVector(cashFlowDates)), rateIndex, new QL.DoubleVector(gearingVector), new QL.DoubleVector(spreadVector), dayCountConvention);
        nonstandardSwap.setPricingEngine(discountingSwapEngine);
        var dataObjectController = DataObjectController.Instance;
        return dataObjectController.Add(handle, nonstandardSwap);

    }


    [ExcelFunction(
    Name = "d.IR_NonStandardInterestRateSwap_GetPrice",
    Description = "Get pricing elements of a Non Standard Interest Rate Swap.",
    Category = "∂Excel: Interest Rates")]
    public static object NonStandardInterestRateSwap_GetPrice(string handle, bool cashFlows = false)
    {
        var dataObjectController = DataObjectController.Instance;
        object dataObject = dataObjectController.GetDataObject(handle);
        object output;


        if (dataObject.GetType() == typeof(QL.NonstandardSwap))
        {
            QL.NonstandardSwap nonstandardSwap = (QL.NonstandardSwap)dataObject;
            if (cashFlows)
            {
                object[,] outputTemp = new object[nonstandardSwap.fixedLeg().Count, 20];
                QL.Date tempFixingDate;
                QL.Date tempFixingStartDate;
                for (int i = 0; i < nonstandardSwap.fixedLeg().Count; i++)
                {

                    //first set of days relate to accrual period. this should be same for fixed and float legs, used fix leg (no specific reason)  to access the data
                    outputTemp[i, 0] = QL.NQuantLibc.as_fixed_rate_coupon(nonstandardSwap.fixedLeg()[i]).accrualStartDate().ToDateTime().ToOADate();
                    outputTemp[i, 1] = QL.NQuantLibc.as_fixed_rate_coupon(nonstandardSwap.fixedLeg()[i]).accrualEndDate().ToDateTime().ToOADate();
                    outputTemp[i, 2] = QL.NQuantLibc.as_fixed_rate_coupon(nonstandardSwap.fixedLeg()[i]).accrualDays();

                    //forecast period can be different from accrual, access this through floating leg

                    tempFixingDate = QL.NQuantLibc.as_floating_rate_coupon(nonstandardSwap.floatingLeg()[i]).fixingDate();
                    outputTemp[i, 3] = tempFixingDate.ToDateTime().ToOADate();                                                                   //actual fixing date
                    tempFixingStartDate = QL.NQuantLibc.as_floating_rate_coupon(nonstandardSwap.floatingLeg()[0]).index().valueDate(tempFixingDate);
                    outputTemp[i, 4] = tempFixingStartDate.ToDateTime().ToOADate();                                                             //traditional forecast start date, i.e. start of underlying NCD                    
                    outputTemp[i, 5] = QL.NQuantLibc.as_floating_rate_coupon(nonstandardSwap.floatingLeg()[0]).index().maturityDate(tempFixingStartDate).ToDateTime().ToOADate(); //traditional forecast end date, i.e. maturity date of underlying NCD

                    outputTemp[i, 6] = (QL.NQuantLibc.as_floating_rate_coupon(nonstandardSwap.floatingLeg()[i]).index().fixingCalendar()).adjust(QL.NQuantLibc.as_fixed_rate_coupon(nonstandardSwap.fixedLeg()[i]).accrualStartDate()).ToOaDate(); //actual forecast start
                    outputTemp[i, 7] = (QL.NQuantLibc.as_floating_rate_coupon(nonstandardSwap.floatingLeg()[i]).index().fixingCalendar()).adjust(QL.NQuantLibc.as_fixed_rate_coupon(nonstandardSwap.fixedLeg()[i]).accrualEndDate()).ToOaDate(); //actual forecast end

                    //payment date. this should be same for fixed and float legs, used fix leg (no specific reason)  to access the data
                    outputTemp[i, 8] = nonstandardSwap.fixedLeg()[i].date().ToDateTime().ToOADate();

                    outputTemp[i, 9] = nonstandardSwap.fixedNominal()[i];
                    outputTemp[i, 10] = QL.NQuantLibc.as_fixed_rate_coupon(nonstandardSwap.fixedLeg()[i]).rate(); //fixed rate
                    outputTemp[i, 11] = nonstandardSwap.fixedLeg()[i].amount(); //fixed projected cf

                    outputTemp[i, 12] = nonstandardSwap.floatingNominal()[i];
                    outputTemp[i, 13] = QL.NQuantLibc.as_floating_rate_coupon(nonstandardSwap.floatingLeg()[i]).index().fixing(tempFixingDate); //floating rate over NCD period
                    outputTemp[i, 14] = QL.NQuantLibc.as_floating_rate_coupon(nonstandardSwap.floatingLeg()[i]).adjustedFixing(); //adjusting floating rate
                    outputTemp[i, 15] = QL.NQuantLibc.as_floating_rate_coupon(nonstandardSwap.floatingLeg()[i]).spread(); //spread over floating rate
                    outputTemp[i, 16] = QL.NQuantLibc.as_floating_rate_coupon(nonstandardSwap.floatingLeg()[i]).gearing();
                    outputTemp[i, 17] = QL.NQuantLibc.as_floating_rate_coupon(nonstandardSwap.floatingLeg()[i]).convexityAdjustment();
                    outputTemp[i, 18] = QL.NQuantLibc.as_floating_rate_coupon(nonstandardSwap.floatingLeg()[i]).rate(); //total rate
                    outputTemp[i, 19] = QL.NQuantLibc.as_floating_rate_coupon(nonstandardSwap.floatingLeg()[i]).amount(); //float projected CF

                }
                output = outputTemp;
            }

            else
            {
                output = nonstandardSwap.NPV();
            }
        }

        else
        {
            return CommonUtils.DExcelErrorMessage("Unknown interest rate swap type.");
        }

        return output;
    }


    /// <summary>
    /// <param name="handle"></param>
    /// <param name="valuationDate">The valuation date.</param>
    /// <param name="effectiveDate">Start date.</param>
    /// <param name="terminationDate">Unadjusted maturity date.</param>
    /// <param name="tenor">Payment/receive frequency (assumed to be the same).</param>
    /// <param name="calendarsToParse">The calendar for the option e.g., 'South Africa' or 'ZAR'.</param>
    /// <param name="businessDayConventionToParse">'Business day convention e.g., 'FOL', 'MODFOL', 'PREC' etc.'.</param>
    /// <param name="ruleToParse">The date generation rule.</param>
    /// <param name="swapTypeToParse">'Payer' or 'Receiver'.</param>
    /// <param name="nominal">Nominal of the swap.</param>
    /// <param name="dayCountConventionToParse">Day count convention e.g., 'Act360' or 'Act365'.</param>
    /// <param name="referenceRateIndexHandle1">Reference Floating Rate Index1.</param>
    /// <param name="referenceRateIndexHandle2">Reference Floating Rate Index2.</param>
    /// <param name="spread1">Spread above floating rate.</param>
    /// <param name="curveHandleDiscount">Handle of curve that will be used for discounting.</param>
    /// <returns></returns>
    [ExcelFunction(
        Name = "d.IR_CreateFloatForFloatInterestRateSwap",
        Description = "Creates a float-for-float interest rate swap object in Excel",
        Category = "∂Excel: Interest Rates")]

    public static string CreateFloatForFloatInterestRateSwap(

        [ExcelArgument(Name = "Handle", Description = DescriptionUtils.Handle)]
        string handle,
        [ExcelArgument(Name = "Valuation Date", Description = "The valuation date.")]
        DateTime valuationDate,
        [ExcelArgument(Name = "Effective Date", Description = "Start date.")]
        DateTime effectiveDate,
        [ExcelArgument(Name = "Termination Date", Description = "Unadjusted maturity date.")]
        DateTime terminationDate,
        [ExcelArgument(Name = "Tenor", Description = "Frequency of cash flows")]
        string tenor,
        [ExcelArgument(Name = "Calendar", Description = "The calendar for the option e.g., 'South Africa' or 'ZAR'.")]
        string calendarsToParse,
        [ExcelArgument(Name = "Business Day Convention", Description = "Business day convention e.g., 'FOL', 'MODFOL', 'PREC' etc.")]
        string businessDayConventionToParse,
        [ExcelArgument(
            Name = "Date generation rule",
            Description = "The date generation rule. " +
                            "\n'Backward' = Start from end date and move backwards. " +
                            "\n'Forward' = Start from start date and move forwards. " +
                            "\n'IMM' = IMM dates.")]
        string ruleToParse,
        [ExcelArgument(Name = "Swap Type", Description = "'Payer' or 'Receiver' (relative to 1st leg). ")]
        string swapTypeToParse,
        [ExcelArgument(Name = "Nominal", Description = "Nominal of the swap.")]
        double nominal,
        [ExcelArgument(Name = "Day Count Convention", Description = "Day count convention e.g., 'Act360' or 'Act365'.")]
        string dayCountConventionToParse,
        [ExcelArgument(Name = "Handle of Floating Rate Index 1", Description = "Reference Floating Rate Index 1.")]
        string referenceRateIndexHandle1,
        [ExcelArgument(Name = "Handle of Floating Rate Index 2", Description = "Reference Floating Rate Index 2.")]
        string referenceRateIndexHandle2,
        [ExcelArgument(Name = "Spread1", Description = "Spread above floating rate index 1.")]
        double spread1,
        [ExcelArgument(Name = "CurveHandleDiscount", Description = "Handle of curve that will be used for discounting.")]
        string curveHandleDiscount)
    {

        Settings.instance().setEvaluationDate(valuationDate.ToQuantLibDate());

        var dataObjectControllerReferenceRateIndex1 = DataObjectController.Instance;
        object dataObjectIndex1 = dataObjectControllerReferenceRateIndex1.GetDataObject(referenceRateIndexHandle1);
        QL.IborIndex rateIndex1 = (QL.IborIndex)dataObjectIndex1;
        var dataObjectControllerReferenceRateIndex2 = DataObjectController.Instance;
        object dataObjectIndex2 = dataObjectControllerReferenceRateIndex2.GetDataObject(referenceRateIndexHandle2);
        QL.IborIndex rateIndex2 = (QL.IborIndex)dataObjectIndex2;

        QL.YieldTermStructure? interestRateCurveDiscount = CurveUtils.GetCurveObject(curveHandleDiscount);

        (QL.Calendar? calendar, string calendarErrorMessage) = DateUtils.ParseCalendars(calendarsToParse);
        (QL.BusinessDayConvention? businessDayConvention, string errorMessage) = DateUtils.ParseBusinessDayConvention(businessDayConventionToParse);
        QL.DayCounter? dayCountConvention = DateUtils.ParseDayCountConvention(dayCountConventionToParse);


        QL.DateGeneration.Rule rule = ruleToParse.ToUpper() switch
        {
            "BACKWARD" => QL.DateGeneration.Rule.Backward,
            "FORWARD" => QL.DateGeneration.Rule.Forward,
            "IMM" => QL.DateGeneration.Rule.TwentiethIMM,
            _ => QL.DateGeneration.Rule.Forward,
        };

        QL.Swap.Type type = swapTypeToParse.ToUpper() switch
        {
            "PAYER" => QL.Swap.Type.Payer,
            "RECEIVER" => QL.Swap.Type.Receiver,
        };


        QL.Date effDateConverted = effectiveDate.ToQuantLibDate();
        QL.Date termDateConverted = terminationDate.ToQuantLibDate();

        QL.Schedule schedule = new(effDateConverted, termDateConverted, new QL.Period(tenor), calendar, (QL.BusinessDayConvention)businessDayConvention, (QL.BusinessDayConvention)businessDayConvention, rule, false);

        //converts nominal into double vector

        List<double> nominalVector = new();
        List<double> spread1Vector = new();
        List<double?> gearingVector = new();

        for (int i = 0; i < schedule.size()-1; i++)
            {
                nominalVector.Add(nominal);
                spread1Vector.Add(spread1);
                gearingVector.Add(1);
        }



        QL.FloatFloatSwap floatFloatSwap = new(type, new QL.DoubleVector(nominalVector), new QL.DoubleVector(nominalVector), schedule, rateIndex1, dayCountConvention, schedule, rateIndex2, dayCountConvention, false, false, new QL.DoubleVector(gearingVector), new QL.DoubleVector(spread1Vector));
   

        QL.DiscountingSwapEngine discountingSwapEngine = new(new QL.YieldTermStructureHandle(interestRateCurveDiscount));
        floatFloatSwap.setPricingEngine(discountingSwapEngine);
        
        var dataObjectController = DataObjectController.Instance;
        return dataObjectController.Add(handle, floatFloatSwap);


    }

    [ExcelFunction(
    Name = "d.IR_FloatForFloatInterestRateSwap_GetPrice",
    Description = "Get pricing elements of a Float-for-Float Interest Rate Swap.",
    Category = "∂Excel: Interest Rates")]
    public static object FloatForFloatInterestRateSwap_GetPrice(string handle, bool cashFlows = false)
    {
        var dataObjectController = DataObjectController.Instance;
        object dataObject = dataObjectController.GetDataObject(handle);
        object output;


        if (dataObject.GetType() == typeof(QL.FloatFloatSwap))
        {
            QL.FloatFloatSwap floatFloatSwap = (QL.FloatFloatSwap)dataObject;
            if (cashFlows)
            {
                object[,] outputTemp = new object[floatFloatSwap.leg(0).Count, 30];
                QL.Date tempFixingDateIndex1;
                QL.Date tempFixingStartDateIndex1;
                QL.Date tempFixingDateIndex2;
                QL.Date tempFixingStartDateIndex2;

                for (int i = 0; i < floatFloatSwap.leg(0).Count; i++)
                {

                    //first set of days relate to accrual period. this should be same for fixed and float legs, used fix leg (no specific reason)  to access the data
                    outputTemp[i, 0] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(0)[i]).accrualStartDate().ToDateTime().ToOADate();
                    outputTemp[i, 1] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(0)[i]).accrualEndDate().ToDateTime().ToOADate();
                    outputTemp[i, 2] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(0)[i]).accrualDays();

                    //forecast period can be different from accrual, access this through each of the floating legs

                    tempFixingDateIndex1 = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(0)[i]).fixingDate();
                    outputTemp[i, 3] = tempFixingDateIndex1.ToDateTime().ToOADate();                                                                   //actual fixing date
                    tempFixingStartDateIndex1 = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(0)[i]).index().valueDate(tempFixingDateIndex1);
                    outputTemp[i, 4] = tempFixingStartDateIndex1.ToDateTime().ToOADate();                                                             //traditional forecast start date, i.e. start of underlying NCD                    
                    outputTemp[i, 5] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(0)[i]).index().maturityDate(tempFixingStartDateIndex1).ToDateTime().ToOADate(); //traditional forecast end date, i.e. maturity date of underlying NCD

                    outputTemp[i, 6] = (QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(0)[i]).index().fixingCalendar()).adjust(QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(0)[i]).accrualStartDate()).ToOaDate(); //actual forecast start
                    outputTemp[i, 7] = (QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(0)[i]).index().fixingCalendar()).adjust(QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(0)[i]).accrualEndDate()).ToOaDate(); //actual forecast end

                    //payment date. this should be same for fixed and float legs, used fix leg (no specific reason)  to access the data
                    outputTemp[i, 8] = floatFloatSwap.leg(0)[i].date().ToDateTime().ToOADate();

                    //floating leg 1
                    outputTemp[i, 9] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(0)[i]).nominal();
                    outputTemp[i, 10] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(0)[i]).index().fixing(tempFixingDateIndex1); //floating rate over NCD period
                    outputTemp[i, 11] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(0)[i]).adjustedFixing(); //adjusting floating rate
                    outputTemp[i, 12] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(0)[i]).spread(); //spread over floating rate
                    outputTemp[i, 13] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(0)[i]).gearing();
                    outputTemp[i, 14] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(0)[i]).convexityAdjustment();
                    outputTemp[i, 15] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(0)[i]).rate(); //total rate
                    outputTemp[i, 16] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(0)[i]).amount(); //float projected CF

                    //floating leg 2

                    tempFixingDateIndex2 = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(1)[i]).fixingDate();
                    outputTemp[i, 17] = tempFixingDateIndex2.ToDateTime().ToOADate();                                                                   //actual fixing date
                    tempFixingStartDateIndex2 = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(1)[i]).index().valueDate(tempFixingDateIndex2);
                    outputTemp[i, 18] = tempFixingStartDateIndex2.ToDateTime().ToOADate();                                                             //traditional forecast start date, i.e. start of underlying NCD                    
                    outputTemp[i, 19] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(1)[i]).index().maturityDate(tempFixingStartDateIndex2).ToDateTime().ToOADate(); //traditional forecast end date, i.e. maturity date of underlying NCD

                    outputTemp[i, 20] = (QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(1)[i]).index().fixingCalendar()).adjust(QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(1)[i]).accrualStartDate()).ToOaDate(); //actual forecast start
                    outputTemp[i, 21] = (QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(1)[i]).index().fixingCalendar()).adjust(QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(1)[i]).accrualEndDate()).ToOaDate(); //actual forecast end


                    outputTemp[i, 22] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(1)[i]).nominal();
                    outputTemp[i, 23] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(1)[i]).index().fixing(tempFixingDateIndex2); //floating rate over NCD period
                    outputTemp[i, 24] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(1)[i]).adjustedFixing(); //adjusting floating rate
                    outputTemp[i, 25] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(1)[i]).spread(); //spread over floating rate
                    outputTemp[i, 26] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(1)[i]).gearing();
                    outputTemp[i, 27] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(1)[i]).convexityAdjustment();
                    outputTemp[i, 28] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(1)[i]).rate(); //total rate
                    outputTemp[i, 29] = QL.NQuantLibc.as_floating_rate_coupon(floatFloatSwap.leg(1)[i]).amount(); //float projected CF

                }
                output = outputTemp;
            }

            else
            {
                output = floatFloatSwap.NPV();
            }
        }

        else
        {
            return CommonUtils.DExcelErrorMessage("Unknown interest rate swap type.");
        }

        return output;
    }

    //non-Excel function to create constant notional ibor leg - this code was repeating in other sections so rather moved to new function
    public static QL.Leg CreateConstantNotionalIborLeg(string handle, DateTime valuationDate, DateTime effectiveDate, DateTime terminationDate, string tenor,
    string calendarsToParse, string businessDayConventionToParse, string ruleToParse, double nominal, string dayCountConventionToParse,
    string referenceRateIndexHandle, double spread, bool includeNotionalExchange)
    {


        var dataObjectControllerReferenceRateIndex = DataObjectController.Instance;
        object dataObjectIndex = dataObjectControllerReferenceRateIndex.GetDataObject(referenceRateIndexHandle);
        QL.IborIndex rateIndex = (QL.IborIndex)dataObjectIndex;


        //first start off through creating a cash flow schedule

        (QL.Calendar? calendar, string calendarErrorMessage) = DateUtils.ParseCalendars(calendarsToParse);
        (QL.BusinessDayConvention? businessDayConvention, string errorMessage) = DateUtils.ParseBusinessDayConvention(businessDayConventionToParse);
        QL.DayCounter? dayCountConvention = DateUtils.ParseDayCountConvention(dayCountConventionToParse);

        QL.DateGeneration.Rule rule = ruleToParse.ToUpper() switch
        {
            "BACKWARD" => QL.DateGeneration.Rule.Backward,
            "FORWARD" => QL.DateGeneration.Rule.Forward,
            "IMM" => QL.DateGeneration.Rule.TwentiethIMM,
            _ => QL.DateGeneration.Rule.Forward,
        };

        QL.Date effDateConverted = effectiveDate.ToQuantLibDate();
        QL.Date termDateConverted = terminationDate.ToQuantLibDate();

        QL.Schedule schedule = new(effDateConverted, termDateConverted, new QL.Period(tenor), calendar, (QL.BusinessDayConvention)businessDayConvention, (QL.BusinessDayConvention)businessDayConvention, rule, false);


        //converts nominal into double vector

        List<double> nominalVector = new();
        List<double> spreadVector = new();
        List<uint> fixingDaysVector = new();

        List<double?> gearingVector = new();

        for (int i = 0; i < schedule.size() - 1; i++)
        {
            //nominalVector.Add(nominal*fxSpot);
            nominalVector.Add(nominal);
            spreadVector.Add(spread);
            fixingDaysVector.Add(rateIndex.fixingDays());
            gearingVector.Add(1);

        }

        QL.Leg leg = QL.NQuantLibc.IborLeg(new QL.DoubleVector(nominalVector), schedule, rateIndex, dayCountConvention, (QL.BusinessDayConvention)businessDayConvention, new QL.UnsignedIntVector(fixingDaysVector), new QL.DoubleVector(gearingVector), new QL.DoubleVector(spreadVector));

        if (includeNotionalExchange)
        {
            leg.Add(new QL.Redemption(nominal, effDateConverted));
            leg.Add(new QL.Redemption(-nominal, termDateConverted));

        }


        return leg;

    }


    /// <summary>
    /// <param name="handle"></param>
    /// <param name="valuationDate">The valuation date.</param>
    /// <param name="effectiveDate">Start date.</param>
    /// <param name="terminationDate">Unadjusted maturity date.</param>
    /// <param name="tenor">Payment/receive frequency (assumed to be the same).</param>
    /// <param name="calendarsToParse">The calendar for the option e.g., 'South Africa' or 'ZAR'.</param>
    /// <param name="businessDayConventionToParse">'Business day convention e.g., 'FOL', 'MODFOL', 'PREC' etc.'.</param>
    /// <param name="ruleToParse">The date generation rule.</param>
    /// <param name="nominal">Nominal of the swap.</param>
    /// <param name="dayCountConventionToParse">Day count convention e.g., 'Act360' or 'Act365'.</param>
    /// <param name="referenceRateIndexHandle">Reference Floating Rate Index.</param>
    /// <param name="spread">Spread above floating rate.</param>
    /// <param name="curveHandleDiscount">Double vector of historical fixings.</param>
    /// <param name="includeNotionalExchange">Exchange notionals indicator.</param>
    /// <returns></returns>
    [ExcelFunction(
        Name = "d.IR_CreateConstantNotionalFloatingLeg",
        Description = "Creates a constant notional cross currency swap object in Excel",
        Category = "∂Excel: Interest Rates")]

    public static string CreateConstantNotionalFloatingLeg(

        [ExcelArgument(Name = "Handle", Description = DescriptionUtils.Handle)]
        string handle,
        [ExcelArgument(Name = "Valuation Date", Description = "The valuation date.")]
        DateTime valuationDate,
        [ExcelArgument(Name = "Effective Date", Description = "Start date.")]
        DateTime effectiveDate,
        [ExcelArgument(Name = "Termination Date", Description = "Unadjusted maturity date.")]
        DateTime terminationDate,
        [ExcelArgument(Name = "Tenor", Description = "Frequency of cash flows")]
        string tenor,
        [ExcelArgument(Name = "Calendar", Description = "The calendar for the option e.g., 'South Africa' or 'ZAR'.")]
        string calendarsToParse,
        [ExcelArgument(Name = "Business Day Convention", Description = "Business day convention e.g., 'FOL', 'MODFOL', 'PREC' etc.")]
        string businessDayConventionToParse,
        [ExcelArgument(
            Name = "Date generation rule",
            Description = "The date generation rule. " +
                            "\n'Backward' = Start from end date and move backwards. " +
                            "\n'Forward' = Start from start date and move forwards. " +
                            "\n'IMM' = IMM dates.")]
        string ruleToParse,
        [ExcelArgument(Name = "Nominal", Description = "Nominal.")]
        double nominal,
        [ExcelArgument(Name = "Day Count Convention", Description = "Day count convention e.g., 'Act360' or 'Act365'.")]
        string dayCountConventionToParse,
        [ExcelArgument(Name = "Handle of Floating Rate Index", Description = "Reference Floating Rate Index.")]
        string referenceRateIndexHandle,
        [ExcelArgument(Name = "Spread", Description = "Spread above floating rate index.")]
        double spread,
        [ExcelArgument(Name = "CurveHandleDiscount", Description = "Handle of curve that will be used for discounting.")]
        string curveHandleDiscount,
        [ExcelArgument(Name = "NotionalExchange", Description = "'True' or 'False' ")]
        bool includeNotionalExchange
        )
    {

        
        Settings.instance().setEvaluationDate(valuationDate.ToQuantLibDate());

        QL.YieldTermStructure? interestRateCurveDiscount = CurveUtils.GetCurveObject(curveHandleDiscount);
        QL.YieldTermStructureHandle yieldTermStructureHandle = new(interestRateCurveDiscount);

        QL.Leg leg = CreateConstantNotionalIborLeg(handle, valuationDate, effectiveDate, terminationDate, tenor,
        calendarsToParse, businessDayConventionToParse, ruleToParse, nominal, dayCountConventionToParse,
        referenceRateIndexHandle, spread, includeNotionalExchange);

        double pv = QL.CashFlows.npv(leg, yieldTermStructureHandle,false);


        object[,] container = new object[2, 1];
        container[0, 0] = leg;
        container[1, 0] = pv;

        var dataObjectController = DataObjectController.Instance;
        return dataObjectController.Add(handle, container);
            

    }


    
    [ExcelFunction(
    Name = "d.IR_ConstantNotionalFloatingLeg_GetPrice",
    Description = "Get pricing elements of a Constant Notional Floating Leg.",
    Category = "∂Excel: Interest Rates")]
    public static object ConstantNotionalFloatingLeg_GetPrice(string handle, bool cashFlows = false)
    {
        var dataObjectController = DataObjectController.Instance;
        object dataObjectContainer = dataObjectController.GetDataObject(handle);
        object output;
        object[,] container = (object[,])dataObjectContainer;
        object dataObject = container[0, 0];
        double pv = (double)container[1, 0];


        if (dataObject.GetType() == typeof(QL.Leg))
        {
            QL.Leg leg = (QL.Leg)dataObject;
            
            if (cashFlows)
            {
                
                    object[,] outputTemp = new object[leg.Count, 17];
                    QL.Date tempFixingDateIndex1;
                    QL.Date tempFixingStartDateIndex1;

                    for (int i = 0; i < leg.Count; i++)
                    {
                    if (QL.NQuantLibc.as_floating_rate_coupon(leg[i]) is not null)
                    {
                        //first set of days relate to accrual period. this should be same for fixed and float legs, used fix leg (no specific reason)  to access the data
                        outputTemp[i, 0] = QL.NQuantLibc.as_floating_rate_coupon(leg[i]).accrualStartDate().ToDateTime().ToOADate();
                        outputTemp[i, 1] = QL.NQuantLibc.as_floating_rate_coupon(leg[i]).accrualEndDate().ToDateTime().ToOADate();
                        outputTemp[i, 2] = QL.NQuantLibc.as_floating_rate_coupon(leg[i]).accrualDays();

                        //forecast period can be different from accrual, access this through each of the floating legs

                        tempFixingDateIndex1 = QL.NQuantLibc.as_floating_rate_coupon(leg[i]).fixingDate();
                        outputTemp[i, 3] = tempFixingDateIndex1.ToDateTime().ToOADate();                                                                   //actual fixing date
                        tempFixingStartDateIndex1 = QL.NQuantLibc.as_floating_rate_coupon(leg[i]).index().valueDate(tempFixingDateIndex1);
                        outputTemp[i, 4] = tempFixingStartDateIndex1.ToDateTime().ToOADate();                                                             //traditional forecast start date, i.e. start of underlying NCD                    
                        outputTemp[i, 5] = QL.NQuantLibc.as_floating_rate_coupon(leg[i]).index().maturityDate(tempFixingStartDateIndex1).ToDateTime().ToOADate(); //traditional forecast end date, i.e. maturity date of underlying NCD

                        outputTemp[i, 6] = (QL.NQuantLibc.as_floating_rate_coupon(leg[i]).index().fixingCalendar()).adjust(QL.NQuantLibc.as_floating_rate_coupon(leg[i]).accrualStartDate()).ToOaDate(); //actual forecast start
                        outputTemp[i, 7] = (QL.NQuantLibc.as_floating_rate_coupon(leg[i]).index().fixingCalendar()).adjust(QL.NQuantLibc.as_floating_rate_coupon(leg[i]).accrualEndDate()).ToOaDate(); //actual forecast end

                        //payment date. this should be same for fixed and float legs, used fix leg (no specific reason)  to access the data
                        outputTemp[i, 8] = leg[i].date().ToDateTime().ToOADate();

                        //floating leg 1
                        outputTemp[i, 9] = QL.NQuantLibc.as_floating_rate_coupon(leg[i]).nominal();
                        outputTemp[i, 10] = QL.NQuantLibc.as_floating_rate_coupon(leg[i]).index().fixing(tempFixingDateIndex1); //floating rate over NCD period
                        outputTemp[i, 11] = QL.NQuantLibc.as_floating_rate_coupon(leg[i]).adjustedFixing(); //adjusting floating rate
                        outputTemp[i, 12] = QL.NQuantLibc.as_floating_rate_coupon(leg[i]).spread(); //spread over floating rate
                        outputTemp[i, 13] = QL.NQuantLibc.as_floating_rate_coupon(leg[i]).gearing();
                        outputTemp[i, 14] = QL.NQuantLibc.as_floating_rate_coupon(leg[i]).convexityAdjustment();
                        outputTemp[i, 15] = QL.NQuantLibc.as_floating_rate_coupon(leg[i]).rate(); //total rate
                        outputTemp[i, 16] = QL.NQuantLibc.as_floating_rate_coupon(leg[i]).amount(); //float projected CF
                    }
                    else if (QL.NQuantLibc.as_floating_rate_coupon(leg[i]) is null)
                    {
                        outputTemp[i, 8] = leg[i].date().ToDateTime().ToOADate();
                        outputTemp[i, 9] = leg[i].amount();
                        outputTemp[i, 16] = leg[i].amount();
                    }

                    }
                    output = outputTemp;

          
            }
            

            else
            {


                output = pv;
            }

            

        }

        else
        {
            return CommonUtils.DExcelErrorMessage("Unknown interest rate swap type.");
        }

        return output;
    }



    /// <summary>
    /// <param name="generalInstrumentInputs">List of inputs applicable to both legs.</param>
    /// <param name="legInputs">List of inputs on individual leg basis.</param>
    /// <returns></returns>
    [ExcelFunction(
        Name = "d.IR_CreateConstantNotionalCrossCurrencySwap",
        Description = "Creates a constant notional cross currency swap object in Excel",
        Category = "∂Excel: Interest Rates")]

    public static string CreateConstantNotionalCrossCurrencySwap(

        [ExcelArgument(Name = "General Instrument Inputs", Description = DescriptionUtils.Handle)]
        object[,] generalInstrumentInputs,
        [ExcelArgument(Name = "Base Currency Leg Inputs", Description = DescriptionUtils.Handle)]
        object[,] legInputs
        )
    {
        int columnHeaderIndexInstrument = ExcelTableUtils.GetRowIndex(generalInstrumentInputs, "Input Name");

        string? instrumentHandle = ExcelTableUtils.GetTableValue<string>(generalInstrumentInputs, "Value", "Xccy Swap Handle", columnHeaderIndexInstrument);
        DateTime valuationDate = ExcelTableUtils.GetTableValue<DateTime>(generalInstrumentInputs, "Value", "Valuation Date", columnHeaderIndexInstrument);
        DateTime effectiveDate = ExcelTableUtils.GetTableValue<DateTime>(generalInstrumentInputs, "Value", "Effective Date", columnHeaderIndexInstrument);
        DateTime terminationDate = ExcelTableUtils.GetTableValue<DateTime>(generalInstrumentInputs, "Value", "Termination Date", columnHeaderIndexInstrument);
        string? tenor = ExcelTableUtils.GetTableValue<string>(generalInstrumentInputs, "Value", "Tenor", columnHeaderIndexInstrument);
        string? calendarsToParse = ExcelTableUtils.GetTableValue<string>(generalInstrumentInputs, "Value", "Calendar", columnHeaderIndexInstrument);
        string? businessDayConventionToParse = ExcelTableUtils.GetTableValue<string>(generalInstrumentInputs, "Value", "Business Day Convention", columnHeaderIndexInstrument);
        string? ruleToParse = ExcelTableUtils.GetTableValue<string>(generalInstrumentInputs, "Value", "Date Generation Rule", columnHeaderIndexInstrument);
        double fxSpot = ExcelTableUtils.GetTableValue<double>(generalInstrumentInputs, "Value", "FXSpot", columnHeaderIndexInstrument);

        int columnHeaderIndexLegs = ExcelTableUtils.GetRowIndex(legInputs, "Input Name");

        string? currencyBaseCcy = ExcelTableUtils.GetTableValue<string>(legInputs, "Base Ccy Inputs", "Currency", columnHeaderIndexLegs);
        double nominalBaseCcy = ExcelTableUtils.GetTableValue<double>(legInputs, "Base Ccy Inputs", "Nominal", columnHeaderIndexLegs);
        string? dayCountConventionToParseBaseCcy = ExcelTableUtils.GetTableValue<string>(legInputs, "Base Ccy Inputs", "Day Count Convention", columnHeaderIndexLegs);
        string? referenceRateBaseCcy = ExcelTableUtils.GetTableValue<string>(legInputs, "Base Ccy Inputs", "Reference Rate Object", columnHeaderIndexLegs);
        double spreadBaseCcy = ExcelTableUtils.GetTableValue<double>(legInputs, "Base Ccy Inputs", "Spread", columnHeaderIndexLegs);
        string? discountCurveHandleBaseCcy = ExcelTableUtils.GetTableValue<string>(legInputs, "Base Ccy Inputs", "Discount Curve Object", columnHeaderIndexLegs);

        string? currencyQuoteCcy = ExcelTableUtils.GetTableValue<string>(legInputs, "Quote Ccy Inputs", "Currency", columnHeaderIndexLegs);
        double nominalQuoteCcy = ExcelTableUtils.GetTableValue<double>(legInputs, "Quote Ccy Inputs", "Nominal", columnHeaderIndexLegs);
        string? dayCountConventionToParseQuoteCcy = ExcelTableUtils.GetTableValue<string>(legInputs, "Quote Ccy Inputs", "Day Count Convention", columnHeaderIndexLegs);
        string? referenceRateQuoteCcy = ExcelTableUtils.GetTableValue<string>(legInputs, "Quote Ccy Inputs", "Reference Rate Object", columnHeaderIndexLegs);
        double spreadQuoteCcy = ExcelTableUtils.GetTableValue<double>(legInputs, "Quote Ccy Inputs", "Spread", columnHeaderIndexLegs);
        string? discountCurveHandleQuoteCcy = ExcelTableUtils.GetTableValue<string>(legInputs, "Quote Ccy Inputs", "Discount Curve Object", columnHeaderIndexLegs);


        Settings.instance().setEvaluationDate(valuationDate.ToQuantLibDate());


        //set discount curves

        QL.YieldTermStructure? DiscountCurveBaseCcy = CurveUtils.GetCurveObject(discountCurveHandleBaseCcy);
        QL.YieldTermStructureHandle yieldTermStructureHandleBaseCcy = new(DiscountCurveBaseCcy);

        QL.YieldTermStructure? DiscountCurveQuoteCcy = CurveUtils.GetCurveObject(discountCurveHandleQuoteCcy);
        QL.YieldTermStructureHandle yieldTermStructureHandleQuoteCcy = new(DiscountCurveQuoteCcy);

        QL.Leg legBaseCcy = CreateConstantNotionalIborLeg(currencyBaseCcy + "Leg", valuationDate, effectiveDate, terminationDate, tenor,
                            calendarsToParse, businessDayConventionToParse, ruleToParse, nominalBaseCcy, dayCountConventionToParseBaseCcy,
                            referenceRateBaseCcy, spreadBaseCcy, true);
        QL.Leg legQuoteCcy = CreateConstantNotionalIborLeg(currencyQuoteCcy + "Leg", valuationDate, effectiveDate, terminationDate, tenor,
                        calendarsToParse, businessDayConventionToParse, ruleToParse, nominalQuoteCcy, dayCountConventionToParseQuoteCcy,
                        referenceRateQuoteCcy, spreadQuoteCcy, true);

        double pvBaseCcy = QL.CashFlows.npv(legBaseCcy, yieldTermStructureHandleBaseCcy, false) * fxSpot;
        double pvQuoteCcy = QL.CashFlows.npv(legQuoteCcy, yieldTermStructureHandleQuoteCcy, false);


        object[,] container = new object[2, 3];
        container[0, 0] = currencyBaseCcy + "Leg"; container[0, 1] = legBaseCcy; container[0, 2] = pvBaseCcy;
        container[1, 0] = currencyQuoteCcy + "Leg"; container[1, 1] = legQuoteCcy; container[1, 2] = pvQuoteCcy;

        var dataObjectController = DataObjectController.Instance;
        return dataObjectController.Add(instrumentHandle, container);

    }


    [ExcelFunction(
    Name = "d.IR_ConstantNotionalXccySwap_GetPrice",
    Description = "Get pricing elements of a Constant Notional XCcy Swap.",
    Category = "∂Excel: Interest Rates")]
    public static object ConstantNotionalXccySwap_GetPrice(string handle, bool cashFlows = false)
    {

        var dataObjectController = DataObjectController.Instance;
        object dataObjectContainer = dataObjectController.GetDataObject(handle);
        object[,] container = (object[,])dataObjectContainer; //input container containing handles, legs and pv

        object output;

        object[,] legs = new object[2, 2];
        legs[0, 0] = container[0,0]; legs[0, 1] = container[0,1]; 
        legs[1, 0] = container[1, 0]; legs[1, 1] = container[1, 1];

        if (legs[0, 1].GetType() == typeof(QL.Leg))
        {

            if (cashFlows)
            {

                QL.Leg legTemp = (QL.Leg)legs[0, 1];
                object[,] outputTemp = new object[legTemp.Count*2, 18];
                int iIndexStart = 0;


                for (int legCounter = 0; legCounter < 2; legCounter++)

                {
                    QL.Leg leg = (QL.Leg)legs[legCounter, 1];


                    QL.Date tempFixingDateIndex1;
                    QL.Date tempFixingStartDateIndex1;

                    int iIndexEnd = leg.Count * (legCounter + 1);
                    int iTempCounter = 0;

                    for (int i = iIndexStart; i < iIndexEnd; i++)
                    {
                        if (QL.NQuantLibc.as_floating_rate_coupon(leg[i-iIndexStart]) is not null)
                        {
                            outputTemp[i, 0] = legs[legCounter, 0];
                            //first set of days relate to accrual period. this should be same for fixed and float legs, used fix leg (no specific reason)  to access the data
                            outputTemp[i, 1] = QL.NQuantLibc.as_floating_rate_coupon(leg[i-iIndexStart]).accrualStartDate().ToDateTime().ToOADate();
                            outputTemp[i, 2] = QL.NQuantLibc.as_floating_rate_coupon(leg[i-iIndexStart]).accrualEndDate().ToDateTime().ToOADate();
                            outputTemp[i, 3] = QL.NQuantLibc.as_floating_rate_coupon(leg[i-iIndexStart]).accrualDays();

                            //forecast period can be different from accrual, access this through each of the floating legs

                            tempFixingDateIndex1 = QL.NQuantLibc.as_floating_rate_coupon(leg[i-iIndexStart]).fixingDate();
                            outputTemp[i, 4] = tempFixingDateIndex1.ToDateTime().ToOADate();                                                                   //actual fixing date
                            tempFixingStartDateIndex1 = QL.NQuantLibc.as_floating_rate_coupon(leg[i-iIndexStart]).index().valueDate(tempFixingDateIndex1);
                            outputTemp[i, 5] = tempFixingStartDateIndex1.ToDateTime().ToOADate();                                                             //traditional forecast start date, i.e. start of underlying NCD                    
                            outputTemp[i, 6] = QL.NQuantLibc.as_floating_rate_coupon(leg[i-iIndexStart]).index().maturityDate(tempFixingStartDateIndex1).ToDateTime().ToOADate(); //traditional forecast end date, i.e. maturity date of underlying NCD

                            outputTemp[i, 7] = (QL.NQuantLibc.as_floating_rate_coupon(leg[i-iIndexStart]).index().fixingCalendar()).adjust(QL.NQuantLibc.as_floating_rate_coupon(leg[i-iIndexStart]).accrualStartDate()).ToOaDate(); //actual forecast start
                            outputTemp[i, 8] = (QL.NQuantLibc.as_floating_rate_coupon(leg[i-iIndexStart]).index().fixingCalendar()).adjust(QL.NQuantLibc.as_floating_rate_coupon(leg[i-iIndexStart]).accrualEndDate()).ToOaDate(); //actual forecast end

                            //payment date. this should be same for fixed and float legs, used fix leg (no specific reason)  to access the data
                            outputTemp[i, 9] = leg[i-iIndexStart].date().ToDateTime().ToOADate();

                            //floating leg 1
                            outputTemp[i, 10] = QL.NQuantLibc.as_floating_rate_coupon(leg[i-iIndexStart]).nominal();
                            outputTemp[i, 11] = QL.NQuantLibc.as_floating_rate_coupon(leg[i-iIndexStart]).index().fixing(tempFixingDateIndex1); //floating rate over NCD period
                            outputTemp[i, 12] = QL.NQuantLibc.as_floating_rate_coupon(leg[i-iIndexStart]).adjustedFixing(); //adjusting floating rate
                            outputTemp[i, 13] = QL.NQuantLibc.as_floating_rate_coupon(leg[i-iIndexStart]).spread(); //spread over floating rate
                            outputTemp[i, 14] = QL.NQuantLibc.as_floating_rate_coupon(leg[i-iIndexStart]).gearing();
                            outputTemp[i, 15] = QL.NQuantLibc.as_floating_rate_coupon(leg[i-iIndexStart]).convexityAdjustment();
                            outputTemp[i, 16] = QL.NQuantLibc.as_floating_rate_coupon(leg[i-iIndexStart]).rate(); //total rate
                            outputTemp[i, 17] = QL.NQuantLibc.as_floating_rate_coupon(leg[i-iIndexStart]).amount(); //float projected CF
                        }
                        else if (QL.NQuantLibc.as_floating_rate_coupon(leg[i-iIndexStart]) is null)
                        {
                            outputTemp[i, 0] = legs[legCounter, 0];
                            outputTemp[i, 9] = leg[i-iIndexStart].date().ToDateTime().ToOADate();
                            outputTemp[i, 10] = leg[i-iIndexStart].amount();
                            outputTemp[i, 17] = leg[i-iIndexStart].amount();
                        }

                        iTempCounter = i;
                    }
                    iIndexStart = iTempCounter+1;

                }

                output = outputTemp;

            }


            else
            {

                output = (double)container[0, 2] + (double)container[1, 2];


            }

        }

        else
        {
            return CommonUtils.DExcelErrorMessage("Unknown interest rate swap type.");
        }
        return output;
    }

}










