using dExcel.CommonEnums;
using NUnit.Framework;
using dExcel.OmicronUtils;
using Omicron;
using Option = Omicron.Option;

namespace dExcelTests.OmicronUtilsTests;

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

    private const string FraJson =
        """
        [{
            "type":
            {
                "$type": "Fra",
                "Tenor": {"amount": 12, "unit": "Month"},
                "ReferenceIndex":
                {
                    "$type": "RateIndex",
                    "Name": "JIBAR",
                    "Tenor": {"amount": 3, "unit": "Month"}
                }
            },
            "date": "2023-02-14T00:00:00",
            "value": 0.07519999999999999 
        }] 
        """;

    private const string FxBasisSwapJson =
        """
        [{
            "type":
            {
                "$type": "FxBasisSwap",
                "BaseIndex":
                {
                    "$type": "RateIndex",
                    "Name": "JIBAR",
                    "Tenor": {"amount": 3, "unit": "Month"}
                },
                "SpreadIndex":
                {
                    "$type": "RateIndex",
                    "Name": "USD-LIBOR",
                    "Tenor": {"amount": 3, "unit": "Month"}
                },
                "Tenor": { "amount": 12, "unit": "Year"}
            },
            "date": "2023-02-14T00:00:00",
            "value": -0.0025
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

    private const string OisJson =
        """
        [{
            "type":
            {
                "$type": "Ois",
                "ReferenceIndex":
                {
                    "$type": "RateIndex",
                    "Name": "FEDFUND",
                    "Tenor": {"amount": 1, "unit": "Day"},
                },
            "Tenor": {"amount": 10, "unit": "Year"},
            },
            "date": "2023-02-14T00:00:00",
            "value": 0.03489
        }]
        """;
    
    [Test]
    public void DeserializeCommodityFutureTest()
    {
        List<QuoteValue>? quoteValues = OmicronUtils.DeserializeOmicronObjects(CommodityFutureJson);
        CommodityFuture commodityFuture = new(new Tenor(12, TenorUnit.Month), Commodity.Ethane);
        Assert.AreEqual(quoteValues?[0].Type, commodityFuture);
    }

    [Test]
    public void DeserializeCommodityOptionTest()
    {
        List<QuoteValue>? quoteValues = OmicronUtils.DeserializeOmicronObjects(CommodityOptionJson);
        CommodityOption commodityOption = new(25, Option.Put, new Tenor(10, TenorUnit.Month), Commodity.BrentCrudeIce);
        Assert.AreEqual(quoteValues?[0].Type, commodityOption);
    }

    [Test]
    public void DeserializeFraForwardTest()
    {
        List<QuoteValue>? quoteValues = OmicronUtils.DeserializeOmicronObjects(FraJson); 
        Fra fra = new(new Tenor(12, TenorUnit.Month), new RateIndex("JIBAR", new Tenor(3, TenorUnit.Month)));
        Assert.AreEqual(quoteValues?[0].Type, fra);
    }

    [Test]
    public void DeserializeFxBasisSwapTest()
    {
        List<QuoteValue>? quoteValues = OmicronUtils.DeserializeOmicronObjects(FxBasisSwapJson); 
        FxBasisSwap fxBasisSwap = 
            new(
                BaseIndex: new RateIndex(RateIndices.JIBAR.ToString(), new Tenor(3, TenorUnit.Month)), 
                SpreadIndex: new RateIndex("USD-LIBOR", new Tenor(3, TenorUnit.Month)), 
                Tenor: new Tenor(12, TenorUnit.Year));
        Assert.AreEqual(quoteValues?[0].Type, fxBasisSwap);
    }
    
    [Test]
    public void DeserializeFxForwardTest()
    {
        List<QuoteValue>? quoteValues = OmicronUtils.DeserializeOmicronObjects(FxForwardJson);
        FxForward fxForward = new(new FxSpot(Currency.ZAR, Currency.USD), new Tenor(1, TenorUnit.Year));
        Assert.AreEqual(quoteValues?[0].Type, fxForward);
    }

    [Test]
    public void DeserializeFxOptionTest()
    {
        List<QuoteValue>? quoteValues = OmicronUtils.DeserializeOmicronObjects(FxOptionJson);
        FxOption fxOption = new(10, new Tenor(2, TenorUnit.Year), new FxSpot(Currency.USD, Currency.ZAR));
        Assert.AreEqual(quoteValues?[0].Type, fxOption);
    }

    [Test]
    public void DeserializeInterestRateSwapTest()
    {
        List<QuoteValue>? quoteValues = OmicronUtils.DeserializeOmicronObjects(InterestRateSwapJson);
        InterestRateSwap interestRateSwap = 
            new(
                ReferenceIndex: new RateIndex("JIBAR", new Tenor(3, TenorUnit.Month)), 
                PaymentFrequency: new Tenor(3, TenorUnit.Month), 
                Tenor: new Tenor(10, TenorUnit.Year)); 
        Assert.AreEqual(quoteValues?[0].Type, interestRateSwap);
    }

    [Test]
    public void DeserializeOisTest()
    {
        List<QuoteValue>? quoteValues = OmicronUtils.DeserializeOmicronObjects(OisJson);     
        Ois ois = new(new RateIndex("FEDFUND", new Tenor(1, TenorUnit.Day)), new Tenor(10, TenorUnit.Year));
        Assert.AreEqual(quoteValues?[0].Type, ois);
    }

    [Test]
    public void DeserializeOmicronRequisition1ExampleTest()
    {
        string rawJson =
            File.ReadAllText(
                @"C:\GitLab\dExcelTools\dExcel\dExcelTests\OmicronUtilsTests\OmicronRequisition1Example.json");
        List<QuoteValue>? quoteValues = OmicronUtils.DeserializeOmicronObjects(rawJson);
        List<QuoteValue> zarSwapCurveQuotes =
            quoteValues
                .Where(x => 
                    (x.Type.GetType() == typeof(RateIndex) && ((RateIndex)x.Type).Name == "JIBAR") ||
                    (x.Type.GetType() == typeof(Fra) && ((Fra)x.Type).ReferenceIndex.Name == "JIBAR") ||
                    (x.Type.GetType() == typeof(InterestRateSwap) && 
                     ((InterestRateSwap)x.Type).ReferenceIndex.Name == "JIBAR"))
                .ToList();
        Assert.AreEqual(zarSwapCurveQuotes.Count, 23);
    }

    [Test]
    public async Task GetSwapCurveQuotesTest()
    {
        string rawJson = 
            await File.ReadAllTextAsync(@"C:\GitLab\dExcelTools\dExcel\dExcelTests\OmicronUtilsTests\OmicronRequisition1Example.json"); 
        List<QuoteValue>? quoteValues = OmicronUtils.DeserializeOmicronObjects(rawJson);
        List<QuoteValue> zarSwapCurveQuotes =
            await OmicronUtils.GetSwapCurveQuotes(RateIndices.JIBAR.ToString(), null, quoteValues);
        Assert.AreEqual(zarSwapCurveQuotes.Count, 23);
    }

    [Test]
    public async Task GetAllSwapCurveQuotesZarSwapTest()
    {
        List<QuoteValue> quotes =
            await OmicronUtils.GetAllSwapCurveQuotes(
                index: RateIndices.JIBAR.ToString(), 
                tenor: new Tenor(Amount: 3, Unit: TenorUnit.Month),
                marketDataDate: new DateTime(year: 2023, month: 03, day: 31));
        Assert.AreEqual(expected: 23, actual: quotes.Count);
    }
    
    [Test]
    public async Task GetAllSwapCurveQuotesUsdSwapTest()
    {
        List<QuoteValue> quotes =
            await OmicronUtils.GetAllSwapCurveQuotes(
                index: RateIndices.USD_LIBOR.ToString(), 
                tenor: new Tenor(3, TenorUnit.Month),
                marketDataDate: new DateTime(2023, 03, 31));
        
        Assert.AreEqual(33, quotes.Count);
    }

    [Test]
    public async Task GetAllFxBasisCurveQuotesTest()
    {
        List<QuoteValue> quotes = 
            await OmicronUtils.GetAllFxBasisCurveQuotes(
                spreadIndexName: RateIndices.USD_LIBOR.ToString(),
                spreadIndexTenor: new Tenor(3, TenorUnit.Month),
                baseIndexName: RateIndices.JIBAR.ToString(),
                baseIndexTenor: new Tenor(3, TenorUnit.Month),
                numeratorCurrency: Currency.ZAR,
                denominatorCurrency: Currency.USD,
                marketDataDate: new DateTime(2023, 03, 31));
        
        Assert.AreEqual(28, quotes.Count);
    }
}
