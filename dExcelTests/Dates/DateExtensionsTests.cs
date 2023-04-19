﻿namespace dExcelTests.Dates;

using dExcel.Dates;
using NUnit.Framework;
using QL = QuantLib;

[TestFixture]
public class DateExtensionsTests
{
    [Test]
    public void ToQuantLibMonthTest()
    {
        DateTime dateTime = new(2023, 03, 31);
        dateTime.Month.ToQuantLibMonth();
    }
    
    [Test]
    public void ToQuantLibDateTest()
    {
        DateTime dateTime = new(2023, 03, 31);
        QL.Date date = dateTime.ToQuantLibDate();
    }
}
