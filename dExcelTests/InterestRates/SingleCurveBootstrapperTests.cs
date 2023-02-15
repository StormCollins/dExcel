namespace dExcelTests.InterestRates;

using dExcel.InterestRates;
using NUnit.Framework;
using QLNet;

[TestFixture]
public class SingleCurveBootstrapperTest
{
    [Test]
    public void BootstrapTest()
    {
        object[,] curveParameters = 
        {
            {"Parameter", "Value"},
            {"BaseDate", 44713},
            {"RateIndexName", "JIBAR"},
            {"Interpolation", "Exponential"},
        };

        object[,] deposits =
        {
            {"Tenors", "Rates", "Include"},
            {"1m", "10%", "TRUE"},
            {"2m", "10%", "TRUE"},
            {"3m", "10%", "TRUE"},
        };

        object[,] fras =
        {
            {"FraTenors", "Rates", "Include"},
            {"3x6", "10%", "TRUE"},
            {"6x9", "10%", "TRUE"},
            {"9x12", "10%", "TRUE"},
        };

        object[,] swaps =
        {
            {"Tenors", "Rates", "Include"},
            {"2y", "10%", "TRUE"},
            {"3y", "10%", "TRUE"},
        };
        
        object[] instruments = {deposits, fras, swaps};
        string handle = SingleCurveBootstrapper.Bootstrap("ZAR-Swap", curveParameters, null, instruments);
    }
    
    
    [Test]
        public void MissingBaseDate()
        {
            object[,] curveParameters =
            {
                {"CurveUtils Parameters", ""},
                {"Parameter", "Value"},
                {"RateIndexName", "JIBAR"},
                {"RateIndexTenor", "3m"},
            };
    
            object[,] instrumentGroups = 
            {
                {"Deposits", "", "", ""},
                {"Tenors", "RateIndex", "Rates", "Include"},
                {"1m", "JIBAR", 0.1, "TRUE"},
                {"3m", "JIBAR", 0.1, "TRUE"},
                {"6m", "JIBAR", 0.1, "TRUE"},
            };
            
            string handle = 
                SingleCurveBootstrapper.Bootstrap(
                    handle: "BootstrappedSingleCurve", 
                    curveParameters: curveParameters,
                    customRateIndex: null,
                    instrumentGroups: instrumentGroups);
            
            const string expected = "#∂Excel Error: Curve parameter missing: 'BASEDATE'.";
            Assert.AreEqual(expected, handle);
        }
        
        [Test]
        public void BootstrapFlatCurveDepositsTest()
        {
            DateTime baseDate = new(2022, 06, 01);
    
            object[,] curveParameters =
            {
                {"CurveUtils Parameters", ""},
                {"Parameter", "Value"},
                {"BaseDate", baseDate.ToOADate()},
                {"RateIndexName", "JIBAR"},
                {"RateIndexTenor", "3m"},
            };
    
            object[,] instruments = 
            {
                {"Deposits", "", "", ""},
                {"Tenors", "RateIndex", "Rates", "Include"},
                {"1m", "JIBAR", 0.1, "TRUE"},
                {"3m", "JIBAR", 0.1, "TRUE"},
                {"6m", "JIBAR", 0.1, "TRUE"},
            };
            
            DayCounter dayCounter = new Actual365Fixed();
            string handle = 
                SingleCurveBootstrapper.Bootstrap(
                    handle: "BootstrappedSingleCurve", 
                    curveParameters, 
                    customRateIndex: null,
                    instruments);

            YieldTermStructure? curve = CurveUtils.GetCurveObject(handle);
            const double tolerance = 0.01; 
            
            Assert.AreEqual(1.0, curve.discount(baseDate));
            Assert.AreEqual(
                expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(1))),
                actual: curve.discount(baseDate.AddMonths(1)),
                delta: tolerance);
            Assert.AreEqual(
                expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(2))),
                actual: curve.discount(baseDate.AddMonths(2)),
                delta: tolerance);
            Assert.AreEqual(
                expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(3))),
                actual: curve.discount(baseDate.AddMonths(3)),
                delta: tolerance);
            Assert.AreEqual(
                expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(6))),
                actual: curve.discount(baseDate.AddMonths(6)),
                delta: tolerance);
        }
        
        [Test]
        public void BootstrapFlatCurveDepositsAndFrasTest()
        {
            DateTime baseDate = new(2022, 06, 01);
            
            object[,] curveParameters =
            {
                {"CurveUtils Parameters", ""},
                {"Parameter", "Value"},
                {"BaseDate", baseDate.ToOADate()},
                {"RateIndexName", "JIBAR"},
                {"RateIndexTenor", "3m"},
            };
    
            object[,] depositInstruments = 
            {
                {"Deposits", "", "", ""},
                {"Tenors", "RateIndex", "Rates", "Include"},
                {"1m", "JIBAR", 0.1, "TRUE"},
                {"3m", "JIBAR", 0.1, "TRUE"},
                {"6m", "JIBAR", 0.1, "TRUE"},
            };
           
            object[,] fraInstruments = 
            {
                {"FRAs", "", "", ""},
                {"FraTenors", "RateIndex", "Rates", "Include"},
                {"6x9", "JIBAR", 0.1, "TRUE"},
                {"9x12", "JIBAR", 0.1, "TRUE"},
            };
    
            object[] instruments = {depositInstruments, fraInstruments};
            Actual365Fixed dayCounter = new();
            string handle = SingleCurveBootstrapper.Bootstrap("BootstrappedSingleCurve", curveParameters, null, instruments);
            YieldTermStructure curve = CurveUtils.GetCurveObject(handle);
            const double tolerance = 0.01; 
            
            Assert.AreEqual(1.0, curve.discount(baseDate));
            Assert.AreEqual(
                expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(1))),
                actual: curve.discount(baseDate.AddMonths(1)),
                delta: tolerance);
            Assert.AreEqual(
                expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(2))),
                actual: curve.discount(baseDate.AddMonths(2)),
                delta: tolerance);
            Assert.AreEqual(
                expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(3))),
                actual: curve.discount(baseDate.AddMonths(3)),
                delta: tolerance);
            Assert.AreEqual(
                expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(6))),
                actual: curve.discount(baseDate.AddMonths(6)),
                delta: tolerance);
            Assert.AreEqual(
                expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(6))) *
                          Math.Exp(-0.1 * dayCounter.yearFraction(baseDate.AddMonths(6), baseDate.AddMonths(9))),
                actual: curve.discount(baseDate.AddMonths(9)),
                delta: tolerance);
        }
        
        [Test]
        public void BootstrapFlatCurveDepositsFrasAndSwapsTest()
        {
            DateTime baseDate = new(2022, 06, 01);
    
            object[,] curveParameters =
            {
                {"CurveUtils Parameters", ""},
                {"Parameter", "Value"},
                {"BaseDate", baseDate.ToOADate()},
                {"RateIndexName", "JIBAR"},
                {"RateIndexTenor", "3m"},
            };
    
            object[,] depositInstruments = 
            {
                {"Deposits", "", "", ""},
                {"Tenors", "RateIndex", "Rates", "Include"},
                {"1m", "JIBAR", 0.1, "TRUE"},
                {"3m", "JIBAR", 0.1, "TRUE"},
                {"6m", "JIBAR", 0.1, "TRUE"},
            };
           
            object[,] fraInstruments = 
            {
                {"FRAs", "", "", ""},
                {"FraTenors", "RateIndex", "Rates", "Include"},
                {"6x9", "JIBAR", 0.1, "TRUE"},
                {"9x12", "JIBAR", 0.1, "TRUE"},
            };
    
            object[,] swapInstruments =
            {
                {"Interest Rate Swaps", "", "", ""},
                {"Tenors", "RateIndex", "Rates", "Include"},
                {"2y", "JIBAR", 0.1, "TRUE"},
                {"3y", "JIBAR", 0.1, "TRUE"},
            };
    
            object[] instruments = {depositInstruments, fraInstruments, swapInstruments};
            Actual365Fixed dayCounter = new();
            string handle = SingleCurveBootstrapper.Bootstrap("BootstrappedSingleCurve", curveParameters, null, instruments);
            
            YieldTermStructure? curve = CurveUtils.GetCurveObject(handle);
            const double tolerance = 0.01; 
            
            Assert.AreEqual(1.0, curve.discount(baseDate));
            Assert.AreEqual(
                expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(1))),
                actual: curve.discount(baseDate.AddMonths(1)),
                delta: tolerance);
            Assert.AreEqual(
                expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(2))),
                actual: curve.discount(baseDate.AddMonths(2)),
                delta: tolerance);
            Assert.AreEqual(
                expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(3))),
                actual: curve.discount(baseDate.AddMonths(3)),
                delta: tolerance);
            Assert.AreEqual(
                expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(6))),
                actual: curve.discount(baseDate.AddMonths(6)),
                delta: tolerance);
            Assert.AreEqual(
                expected: Math.Exp(-0.1 * dayCounter.yearFraction(baseDate, baseDate.AddMonths(6))) *
                          Math.Exp(-0.1 * dayCounter.yearFraction(baseDate.AddMonths(6), baseDate.AddMonths(9))),
                actual: curve.discount(baseDate.AddMonths(9)),
                delta: tolerance);
        }
}
