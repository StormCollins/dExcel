namespace dExcelTests;

using dExcel.Curves;
using dExcel.InterestRates;
using NUnit.Framework;
using QLNet;

[TestFixture]
public class CurveBootstrapperTests
{
    [Test]
    public void UsdOisTest()
    {
        var fedFundsIndex = new FedFunds();
        var settlementDate = new Date(04, 04, 2022);
        Settings.setEvaluationDate(settlementDate);

        var oisOnDeposit =
            new DepositRateHelper(
                new Handle<Quote>(new SimpleQuote(0.00330)),
            new Period("1d"),
                fedFundsIndex.fixingDays(),
                fedFundsIndex.fixingCalendar(),
                fedFundsIndex.businessDayConvention(),
                fedFundsIndex.endOfMonth(),
                fedFundsIndex.dayCounter());
//        USDONFSR = X  0.33171
//USDSWFSR = X	#N/A
//USD1MFSR = X  0.452
//USD2MFSR = X	#N/A
//USD3MFSR = X  0.96157
//USD1X4F = 1.156
//USD2X5F = 1.473
//USD3X6F = 1.666
//USD4X7F = 1.872
//USD5X8F = 2.109
//USD6X9F = 2.2986
//USD7X10F = 2.4846
//USD8X11F = 2.636
//USD9X12F = 2.7651
//USD12X15F = 3.055
//USD15X18F = 3.1827
//USD18X21F = 3.149
//USD21X24F = FMD   3.01
//USDSB3L2Y = 2.5665
//USDSB3L3Y = 2.656
//USDSB3L4Y = 2.599
//USDSB3L5Y = 2.5137
//USDSB3L6Y = 2.476
//USDSB3L7Y = 2.4513
//USDSB3L8Y = 2.454
//USDSB3L9Y = 2.409
//USDSB3L10Y = 2.3951
//USDSB3L12Y = 2.4051
//USDSB3L15Y = 2.42625
//USDSB3L20Y = 2.3693
//USDSB3L25Y = 2.34
//USDSB3L30Y = 2.2424
//USDSB3L40Y = 2.108

        // var oisRateHelper1d = new DatedOISRateHelper(settlementDate, new Date(05, 05, 2022),
        //             new Handle<Quote>(new SimpleQuote(0.00330)), fedFundsIndex);
        var oisRateHelper1m = new DatedOISRateHelper(settlementDate, new Date(04, 05, 2022),
            new Handle<Quote>(new SimpleQuote(0.00342)), fedFundsIndex);
        // var oisRateHelper1m = new OISRateHelper(2, new Period("1M"), new Handle<Quote>(new SimpleQuote(0.00342)), fedFundsIndex);
        var oisRateHelper2m = new DatedOISRateHelper(settlementDate, new Date(06, 06, 2022),
            new Handle<Quote>(new SimpleQuote(0.00565)), fedFundsIndex);
        // var oisRateHelper2m = new OISRateHelper(2, new Period("2M"), new Handle<Quote>(new SimpleQuote(0.00565)), fedFundsIndex);
        var oisRateHelper3m = new DatedOISRateHelper(settlementDate, new Date(05, 07, 2022),
            new Handle<Quote>(new SimpleQuote(0.00718)), fedFundsIndex);
        // var oisRateHelper3m = new OISRateHelper(2, new Period("3M"), new Handle<Quote>(new SimpleQuote(0.00718)), fedFundsIndex);
        var oisRateHelper4m = new DatedOISRateHelper(settlementDate, new Date(04, 08, 2022),
            new Handle<Quote>(new SimpleQuote(0.00856)), fedFundsIndex);
        // var oisRateHelper4m = new OISRateHelper(2, new Period("4M"), new Handle<Quote>(new SimpleQuote(0.00856)), fedFundsIndex);
        var oisRateHelper5m = new OISRateHelper(2, new Period("5M"), new Handle<Quote>(new SimpleQuote(0.01009)), fedFundsIndex);
        var oisRateHelper6m = new OISRateHelper(2, new Period("6M"), new Handle<Quote>(new SimpleQuote(0.01118)), fedFundsIndex);
        var oisRateHelper9m = new OISRateHelper(2, new Period("9M"), new Handle<Quote>(new SimpleQuote(0.01458)), fedFundsIndex);
        var oisRateHelper12m = new OISRateHelper(2, new Period("12M"), new Handle<Quote>(new SimpleQuote(0.01731)), fedFundsIndex);
        var oisRateHelper18m = new OISRateHelper(2, new Period("18M"), new Handle<Quote>(new SimpleQuote(0.02088)), fedFundsIndex);
        var oisRateHelper2y = new OISRateHelper(2, new Period("2y"), new Handle<Quote>(new SimpleQuote(0.02305)), fedFundsIndex);
        var oisRateHelper3y = new OISRateHelper(2, new Period("3y"), new Handle<Quote>(new SimpleQuote(0.02345)), fedFundsIndex);
        var oisRateHelper4y = new OISRateHelper(2, new Period("4y"), new Handle<Quote>(new SimpleQuote(0.02285)), fedFundsIndex);
        var oisRateHelper5y = new OISRateHelper(2, new Period("5y"), new Handle<Quote>(new SimpleQuote(0.02210)), fedFundsIndex);
        var oisRateHelper7y = new OISRateHelper(2, new Period("7y"), new Handle<Quote>(new SimpleQuote(0.02138)), fedFundsIndex);
        var oisRateHelper10y = new OISRateHelper(2, new Period("10y"), new Handle<Quote>(new SimpleQuote(0.02097)), fedFundsIndex);
        var oisRateHelper12y = new OISRateHelper(2, new Period("12y"), new Handle<Quote>(new SimpleQuote(0.02100)), fedFundsIndex);

        var oisHelpers = new List<RateHelper>
        {
            oisOnDeposit,
            oisRateHelper1m,
            oisRateHelper2m,
            oisRateHelper3m,
            oisRateHelper4m,
            oisRateHelper5m,
            oisRateHelper6m,
            oisRateHelper9m,
            oisRateHelper12m,
            oisRateHelper18m,
            oisRateHelper2y,
            oisRateHelper3y,
            oisRateHelper4y,
            oisRateHelper5y,
            oisRateHelper7y,
            oisRateHelper10y,
            oisRateHelper12y,
        };

        var tolerance = 1.0e-14;
        var oisCurve = 
            new PiecewiseYieldCurve<Discount, LogLinear>(
                settlementDate, oisHelpers, fedFundsIndex.dayCounter(), new List<Handle<Quote>>(), new List<Date>(), tolerance);

        Date d1 = new Date(31, 03, 2022);
        Date d2 = new Date(04, 04, 2022);
        Date d3 = new Date(05, 04, 2022);
        Date d4 = new Date(04, 05, 2022);
        Date d5 = new Date(06, 06, 2022);
        Date d6 = new Date(05, 07, 2022);
        Date d7 = new Date(04, 08, 2022);
        Date d8 = new Date(06, 09, 2022);
        Date d9 = new Date(04, 10, 2022);
        Date d10 = new Date(04, 01, 2023);
        Date d11 = new Date(04, 04, 2023);
        Date d12 = new Date(04, 10, 2023);
        Date d13 = new Date(04, 04, 2024);

        // var df1 = oisCurve.discount(d1);
        var df2 = oisCurve.discount(d2);
        var df3 = oisCurve.discount(d3);
        var df4 = oisCurve.discount(d4);
        var df5 = oisCurve.discount(d5);
        var df6 = oisCurve.discount(d6);
        var df7 = oisCurve.discount(d7);
        var df8 = oisCurve.discount(d8);
        var df9 = oisCurve.discount(d9);
        var df10 = oisCurve.discount(d10);
        var df11 = oisCurve.discount(d11);
        var df12 = oisCurve.discount(d12);
        var df13 = oisCurve.discount(d13);

        // ---------------------------------------------------------------
        // Dual CurveUtils bootstrapping 
        RelinkableHandle<YieldTermStructure> forecastCurve = new RelinkableHandle<YieldTermStructure>();
        USDLibor libor = new USDLibor(new Period("3m"), forecastCurve);

        //DepositRateHelper depositRateHelper = new DepositRateHelper()
    }

    [Test]
    public void GetTest()
    {
        string something = CurveBootstrapper.Get("Something", "USD-OIS", new DateTime(2023, 3, 28));
        object[] dates = 
        {
            new DateTime(2024, 03, 28).ToOADate(), 
            new DateTime(2025, 3, 28).ToOADate(), 
        };
        
        var discountFactors = CurveUtils.GetDiscountFactors(something, dates);
        // Assert.AreEqual("USD-OIS", something);
    }
}
