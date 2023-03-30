namespace dExcelTests.OmicronUtilsTests;

using NUnit.Framework;
using dExcel.OmicronUtils;
using Omicron;
using Option = Omicron.Option;

[TestFixture]
public class OmicronUtilsTests
{
    private const string CommodityFutureJson = 
        """
        [{ 
            "type": 
            {
                "$type": "CommodityFuture",
                "Tenor": {"amount":12, "unit": "Month"},
                "Commodity": "Ethane" 
            },
            "date": "2023-02-14T00:00:00",
            "value": 0.24625 
        }]
        """;

    private const string CommodityOptionJson = 
        """
        [{
            "type":
            {
                "$type": "CommodityOption",
                "Delta": 25, 
                "OptionType": "Put",
                "Tenor": {"amount": 10, "unit": "Month"},
                "Commodity": "BrentCrudeIce"
            },
            "date": "2023-02-14T00:00:00",
            "value": 0.4103 
        }]
        """;

    private const string FxForwardJson =
        """
        [{
            "type":
            {
                "$type": "FxForward",
                "FxSpot":
                {
                    "$type": "FxSpot",
                    "Numerator": "ZAR",
                    "Denominator": "USD"
                },
                "Tenor": 
                {
                    "amount": 1,
                    "unit": "Year"
                }
            },
            "date": "2023-02-14T00:00:00",
            "value": 0.46585
        }]
        """;
    
    private const string FxOptionJson =
        """
        [{
            "type":
            {
                "$type": "FxOption",
                "Delta": 10,
                "Tenor": {"amount": 2, "unit": "Year"},
                "ReferenceSpot": 
                {
                    "$type": "FxSpot",
                    "Numerator": "USD",
                    "Denominator": "ZAR"
                }
            },
            "date": "2023-02-14T00:00:00",
            "value": 0.15864
        }]
        """;

    private const string InterestRateSwapJson = 
        """
        [{ 
            "type": 
            { 
                "$type": "InterestRateSwap",
                "ReferenceIndex":
                {
                    "$type": "RateIndex",
                    "Name": "JIBAR",
                    "Tenor": {"amount": 3, "unit": "Month"} 
                },
                "PaymentFrequency": {"amount": 3, "unit": "Month"},
                "Tenor": {"amount": 10, "unit": "Year"}
            },
            "date": "2023-02-14T00:00:00",
            "value": 0.08779999999999999
        }]
        """;

    [Test]
    public void DeserializeCommodityFutureTest()
    {
        List<QuoteValue> quoteValues = OmicronUtils.DeserializeOmicronObject(CommodityFutureJson);
        CommodityFuture commodityFuture = new(new Tenor(12, TenorUnit.Month), Commodity.Ethane);
        Assert.AreEqual(quoteValues[0].Type, commodityFuture);
    }

    [Test]
    public void DeserializeCommodityOptionTest()
    {
        List<QuoteValue> quoteValues = OmicronUtils.DeserializeOmicronObject(CommodityOptionJson);
        CommodityOption commodityOption = new(25, Option.Put, new Tenor(10, TenorUnit.Month), Commodity.BrentCrudeIce);
        Assert.AreEqual(quoteValues[0].Type, commodityOption);
    }

    [Test]
    public void DeserializeFxForwardTest()
    {
        List<QuoteValue> quoteValues = OmicronUtils.DeserializeOmicronObject(FxForwardJson);
        FxForward fxForward = new(new FxSpot(Currency.ZAR, Currency.USD), new Tenor(1, TenorUnit.Year));
        Assert.AreEqual(quoteValues[0].Type, fxForward);
    }

    [Test]
    public void DeserializeFxOptionTest()
    {
        List<QuoteValue> quoteValues = OmicronUtils.DeserializeOmicronObject(FxOptionJson);
        FxOption fxOption = new(10, new Tenor(2, TenorUnit.Year), new FxSpot(Currency.USD, Currency.ZAR));
        Assert.AreEqual(quoteValues[0].Type, fxOption);
    }

    [Test]
    public void DeserializeInterestRateSwapTest()
    {
        List<QuoteValue> quoteValues = OmicronUtils.DeserializeOmicronObject(InterestRateSwapJson);
        InterestRateSwap interestRateSwap = 
            new(
                ReferenceIndex: new RateIndex("JIBAR", new Tenor(3, TenorUnit.Month)), 
                PaymentFrequency: new Tenor(3, TenorUnit.Month), 
                Tenor: new Tenor(10, TenorUnit.Year)); 
        Assert.AreEqual(quoteValues[0].Type, interestRateSwap);
    }
}
