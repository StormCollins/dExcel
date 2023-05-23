using NUnit.Framework;
using QL = QuantLib;

namespace dExcelTests.InterestRates;

public class Instruments
{
    public double B(double S, double T, QL.HullWhiteProcess hullWhiteProcess)
    {
        double alpha = hullWhiteProcess.drift(0, 0);
        return (1 - Math.Exp(-1 * alpha * (T - S))) / alpha;
    }
    
    public double A(double S, double T, QL.HullWhiteProcess hullWhiteProcess, QL.YieldTermStructureHandle initialTermStructure)
    {
        double pT = initialTermStructure.discount(T);
        double pS = initialTermStructure.discount(S);
        double delta = 0.001;
        double pSDeltaT = initialTermStructure.discount(S + delta);
        double derivativeTerm = (Math.Log(pSDeltaT) - Math.Log(pS)) / delta;
        double alpha = hullWhiteProcess.drift(0, 0);
        double sigma = hullWhiteProcess.diffusion(0, 0);
        return (pT / pS) * 
               Math.Exp(
                   -1 * B(S, T, hullWhiteProcess) * derivativeTerm -
                   sigma * sigma * Math.Pow((Math.Exp(-1 * alpha * T) - Math.Exp(-1 * alpha * S)), 2) * 
                   (Math.Exp(2 * alpha * S) - 1)/ (4 * Math.Pow(alpha, 3)));
    }

    public QL.DiscountCurve CreateHullWhiteDiscountCurve(
        double S, 
        double rt,
        QL.Date currentDate,
        QL.HullWhiteProcess hullWhiteProcess,
        QL.YieldTermStructureHandle initialTermStructure,
        QL.DoubleVector timeSteps,
        QL.DateVector dates)
    {
        List<double> relevantTimes = timeSteps.Where(t => t > S).ToList();
        List<QL.Date> relevantDates = dates.Where(d => d > currentDate).ToList();
        List<double> discountFactors = new();
        
        foreach (double t in relevantTimes)
        {
            discountFactors.Add(A(S, t, hullWhiteProcess, initialTermStructure) * Math.Exp(-1 * rt * B(S, t, hullWhiteProcess)));
        }

        relevantTimes = relevantTimes.Select(x => x - S).ToList();
        relevantTimes.Insert(0, 0);
        relevantDates.Insert(0, currentDate);
        discountFactors.Insert(0, 1);
        
        QL.DiscountCurve discountCurve = 
            new(
            new QL.DateVector(relevantDates), 
            new QL.DoubleVector(discountFactors), 
            new QL.Actual365Fixed());
        
        return discountCurve;
    }
    
    [Test]
    public void TestCVA()
    {
        QL.Date referenceDate = new(31, QL.Month.March, 2023);
        QL.YieldTermStructureHandle termStructure =
            new(
                new QL.FlatForward(
                    referenceDate: referenceDate,
                    forward: new QL.QuoteHandle(new QL.SimpleQuote(0.1)),
                    dayCounter: new QL.Actual365Fixed()));
        
        QL.HullWhiteProcess hullWhiteProcess0 = new(termStructure, 0.25, 0.02);
        QL.HullWhiteProcess hullWhiteProcess1 = new(termStructure, 0.15, 0.05);
        QL.StochasticProcess1DVector stochasticProcess1DVector = new();
        stochasticProcess1DVector.Add(hullWhiteProcess0);
        stochasticProcess1DVector.Add(hullWhiteProcess1);
        QL.Matrix correlationMatrix = new(2, 2);
        correlationMatrix.set(0, 0, 1);
        correlationMatrix.set(0, 1, 0.5);
        correlationMatrix.set(1, 0, 0.5);
        correlationMatrix.set(1, 1, 1);
        QL.StochasticProcessArray stochasticProcessArray = new(stochasticProcess1DVector, correlationMatrix);

        int numberOfTimeSteps = 5;
        double maturity = 1.0;
        double timeStepSize = maturity / numberOfTimeSteps;
        QL.Date maturityDate = new(31, QL.Month.March, 2024);
        QL.Calendar calendar = new QL.SouthAfrica();
        QL.Schedule schedule = new(
            effectiveDate: referenceDate, 
            terminationDate: maturityDate, 
            tenor: new QL.Period("3m"), 
            calendar: calendar, 
            convention: QL.BusinessDayConvention.ModifiedFollowing, 
            terminationDateConvention: QL.BusinessDayConvention.ModifiedFollowing,
            rule: QL.DateGeneration.Rule.Forward,
            endOfMonth: false);
       
        QL.DayCounter dayCounter = new QL.Actual365Fixed();
        QL.DoubleVector timeSteps = new(schedule.dates().Select(x => dayCounter.yearFraction(referenceDate, x)));
        QL.TimeGrid timeGrid = new(timeSteps);
        QL.UniformRandomGenerator uniformRandomGenerator = new();
        QL.UniformRandomSequenceGenerator uniformRandomSequenceGenerator = new(2 * 4, uniformRandomGenerator);
        QL.GaussianRandomSequenceGenerator gaussianRandomSequenceGenerator = new(uniformRandomSequenceGenerator);
        QL.GaussianMultiPathGenerator multiPathGenerator = 
            new(stochasticProcessArray, timeGrid, gaussianRandomSequenceGenerator);
        
        QL.MultiPath paths = multiPathGenerator.next().value();

        for (uint hullWhiteProcessPath = 0; hullWhiteProcessPath < 2; hullWhiteProcessPath++)
        {
            Console.Write($"Hull-White Process {hullWhiteProcessPath}: ");
            for (uint timeStep = 0; timeStep < 5; timeStep++)
            {
               Console.Write($"{paths.at(hullWhiteProcessPath).value(timeStep):0.####} "); 
            }

            Console.WriteLine();
        }

        paths = multiPathGenerator.next().value();

        for (uint hullWhiteProcessPath = 0; hullWhiteProcessPath < 2; hullWhiteProcessPath++)
        {
            Console.Write($"Hull-White Process {hullWhiteProcessPath}: ");
            for (uint timeStep = 0; timeStep < 5; timeStep++)
            {
               Console.Write($"{paths.at(hullWhiteProcessPath).value(timeStep):0.####} "); 
            }

            Console.WriteLine();
        }

        
    }
}
