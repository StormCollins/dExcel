namespace dExcelTests.ExcelUtilsTests;

using dExcel.ExcelUtils;
using NUnit.Framework;

[TestFixture]
public class ExcelArrayUtilsTests
{
    [Test]
    public void ConvertExcelRangeToListForAColumnTest()
    {
        object[,] range = {{1.0}, {2.0}, {3.0}, {4.0}};
        List<double> actual = ExcelArrayUtils.ConvertExcelRangeToList<double>(range);
        List<double> expected = new() {1.0, 2.0, 3.0, 4.0};
        Assert.AreEqual(expected, actual);
    }
    
    [Test]
    public void ConvertExcelRangeToListForARowTest()
    {
        object[,] range = {{1.0, 2.0, 3.0, 4.0}};
        List<double> actual = ExcelArrayUtils.ConvertExcelRangeToList<double>(range);
        List<double> expected = new() {1.0, 2.0, 3.0, 4.0};
        Assert.AreEqual(expected, actual);
    }

    [Test]
    public void ConvertExcelRangeOfDatesFromColumnTest()
    {
        object[,] range = 
            {
                {new DateTime(2022, 01, 01).ToOADate()}, 
                {new DateTime(2022, 02, 01).ToOADate()},
                {new DateTime(2022, 03, 01).ToOADate()},
            };

        List<DateTime> expected = new() {new(2022, 01, 01), new(2022, 02, 01), new(2022, 03, 01)};
        List<DateTime> actual = ExcelArrayUtils.ConvertExcelRangeToList<DateTime>(range);
        Assert.AreEqual(expected, actual);
    }
    
    [Test]
    public void ConvertExcelRangeOfDatesFromRowTest()
    {
        object[,] range = 
            { 
                {
                    new DateTime(2022, 01, 01).ToOADate(), 
                    new DateTime(2022, 02, 01).ToOADate(), 
                    new DateTime(2022, 03, 01).ToOADate()
                } 
            };

        List<DateTime> expected = new() {new(2022, 01, 01), new(2022, 02, 01), new(2022, 03, 01)};
        List<DateTime> actual = ExcelArrayUtils.ConvertExcelRangeToList<DateTime>(range);
        Assert.AreEqual(expected, actual);
    }

    [Test]
    public void ConvertListToExcelColumnRangeTest()
    {
        List<double> list = new() {1.0, 2.0, 3.0};
        object[,] expected = {{1.0}, {2.0}, {3.0}};
        object[,] actual = ExcelArrayUtils.ConvertListToExcelRange(list, 0); 
        Assert.AreEqual(expected, actual);
    }

    [Test]
    public void ConvertListToExcelRowRangeTest()
    {
        List<double> list = new() {1.0, 2.0, 3.0};
        object[,] expected = {{1.0, 2.0, 3.0}};
        object[,] actual = ExcelArrayUtils.ConvertListToExcelRange(list, 1);
        Assert.AreEqual(expected, actual);
    }
}
