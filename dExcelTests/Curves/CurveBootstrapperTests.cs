namespace dExcelTests;

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
    }
}
