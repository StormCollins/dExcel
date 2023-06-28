//using dExcel.InterestRates;
//using dExcel.Utilities;
using dExcel.Dates;
using NUnit.Framework;
using NUnit.Framework.Internal;
using Omicron;
using QuantLib;
using System.Globalization;
using QL = QuantLib;

namespace dExcelTests.InterestRates;
//namespace dExcel.Dates;


[TestFixture]

public class hwTest
{
    [Test]
    public void HWRandomNumberGen()
    {
            
        double sigma = 0.1;
        double a = 0.1;
        uint timestep = 500;
        double length = 5; // in years
        double forward_rate = 0.05;
        int sims = 10000;
        double[,] myArrayShortRate = new double[sims,timestep+1];
        double[,] myArrayBondPrices = new double[sims, timestep + 1];
        double[,] myArrayStochDF = new double[sims, timestep + 1];
//        double[,] myArrayCurveSims = new double[sims, 20];

        QL.DayCounter day_count = new QL.Actual360();
            //Thirty360(QL.Thirty360.Convention.BondBasis);//Actual365Fixed();//QL.Thirty360(QL.Thirty360.Convention.BondBasis);
        QL.Date todays_date = new QL.Date(15, 1.ToQuantLibMonth(), 2015);;
        QL.Settings.instance().setEvaluationDate(todays_date);
        QL.FlatForward spot_curve = new(todays_date, new QL.QuoteHandle(new QL.SimpleQuote(forward_rate)), day_count);
        QL.YieldTermStructureHandle yieldTermStructureHandle = new(spot_curve);
        QL.HullWhiteProcess hw_process = new(yieldTermStructureHandle, a, sigma);//HullWhiteProcess(spot_curve_handle, a, sigma)
        QL.GaussianRandomSequenceGenerator rng = new(new QL.UniformRandomSequenceGenerator(timestep, new QL.UniformRandomGenerator()));
        QL.GaussianPathGenerator seq = new(hw_process, length, timestep, rng, false);
        QL.HullWhite hullWhite = new(yieldTermStructureHandle, a, sigma);

        

        for (int i = 0; i < sims; i++)
        {
            double stochDiscFctrSum = 0;
            QL.Path path = seq.next().value();
            for (uint j = 0; j < timestep + 1;j++)
            {

                //attempt to set up curve given simulated short rate
                List<DateTime> curveDatesVector = new();//check data type
                List<double> curveDFVector = new();

                for (int k = 0; k < 30*12; k++)
                {
                    curveDatesVector.Add(todays_date.ToDateTime().AddDays(path.time(j)*360+k*30));
                    curveDFVector.Add(hullWhite.discountBond(path.time(j), (path.time(j) * 360 + k * 30)/360, path.value(j)));
//                    if (j == 1)
//                    {
//                        myArrayCurveSims[i, k] = curveDFVector[k];

//                    };
                }

                QL.DiscountCurve discountCurve = new(new QL.DateVector(curveDatesVector), new QL.DoubleVector(curveDFVector), day_count);
                QL.YieldTermStructureHandle yieldTermStructureHandleUpdated = new(discountCurve);


                myArrayShortRate[i,j] = path.value(j);
                myArrayBondPrices[i, j] = hullWhite.discountBond(path.time(j), length, path.value(j));//Math.Exp(-stochDiscFctrSum)*hullWhite.discountBond(path.time(j), length, path.value(j));
                myArrayStochDF[i, j] = Math.Exp(-stochDiscFctrSum);
                stochDiscFctrSum += path.value(j) * (length/timestep);

            }
        }
//        WriteToFile("C:/temp/hwCurveSims.csv", myArrayCurveSims);
        WriteToFile("C:/temp/hwShortRate.csv", myArrayShortRate);
        WriteToFile("C:/temp/hwBondPrices.csv", myArrayBondPrices);
        WriteToFile("C:/temp/hwStochDF.csv", myArrayStochDF);
    }

 






    public static void WriteToFile(string filename, double[,] values)
    {
        using (var sr = new StreamWriter(filename))
        {
            for (var row = 0; row < values.GetLength(0); row++)
            {
                if (row > 0) sr.Write("\n");
                for (var col = 0; col < values.GetLength(1); col++)
                {
                    if (col > 0) sr.Write(",");
                    sr.Write(values[row, col].ToString(CultureInfo.InvariantCulture));
                }
            }
        }
    }


}