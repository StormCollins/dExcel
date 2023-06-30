using System.Net.Http;
using System.Net.Http.Headers;
using dExcel.Utilities;
using JsonConverter = Newtonsoft.Json.JsonConverter;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Linq;
using Omicron;
using Option = Omicron.Option;

namespace dExcel.OmicronUtils;

/// <summary>
/// A collection of utility functions for interfacing with Omicron, the market data database.
/// </summary>
public static class OmicronUtils
{
    /// <summary>
    /// The URL to Omicron.
    /// </summary>
    private const string OmicronUrl = "https://omicron.fsa-aks.deloitte.co.za";

    /// <summary>
    /// Deserializes Omicron objects from a given JSON string.
    /// </summary>
    /// <param name="json">The JSON string to deserialize.</param>
    /// <returns>A list of deserialized Omicron 'QuoteValues'.</returns>
    public static List<QuoteValue>? DeserializeOmicronObjects(string json)
    {
        JsonConvert.DefaultSettings = () => new JsonSerializerSettings
        {
            Converters = new List<JsonConverter>
            {
                new DExcelQuoteTypeConverter(), 
                new StringEnumConverter(),
                new DExcelQuoteValueConverter(),
            }
        };
        
        return JsonConvert.DeserializeObject<List<QuoteValue>>(json);
    }

    /// <summary>
    /// Extracts the relevant swap curve quotes from a either Omicron directly or a pre-deserialized list of Omicron
    /// 'QuoteValues'. Thus one must either populate the <param name="quotes"/> or populate both the
    /// <param name="requisitionId"/> and .
    /// </summary>
    /// <param name="index">In the case of an IRS, the index name e.g., 'JIBAR'. In the case of a cross currency
    /// swap this is the same as the quote index i.e., 'JIBAR' in the case of 'USDZAR'.</param>
    /// <param name="spreadIndex">This is null in the case of an IRS. For a cross currency swap it's the same as the
    /// base index i.e., for 'USDZAR' it would be 'USD-LIBOR'.</param>
    /// <param name="quotes">(Optional)The list of Omicron quote values to loop through. If the</param>
    /// <param name="requisitionId">The relevant requisition ID in Omicron.</param>
    /// <param name="marketDataDate">The market data date for which to extract the data from Omicron.</param>
    /// <returns>A list of the relevant swap curve quotes.</returns>
    public static async Task<List<QuoteValue>> GetSwapCurveQuotes(
        string index,
        string? spreadIndex = null,
        List<QuoteValue>? quotes = null, 
        int? requisitionId = null,
        string? marketDataDate = null)
    {
        quotes ??= GetOmicronRequisitionData(requisitionId, marketDataDate);
        // Enums don't allow for hyphens but Omicron uses them.
        index = index.Replace("_", "-");
        return 
            quotes
                .Where(x => 
                    (x.Type.GetType() == typeof(RateIndex) && ((RateIndex) x.Type).Name == index) ||
                    (x.Type.GetType() == typeof(Fra) && ((Fra)x.Type).ReferenceIndex.Name == index) ||
                    (x.Type.GetType() == typeof(InterestRateSwap) && 
                     ((InterestRateSwap)x.Type).ReferenceIndex.Name == index) ||
                    (x.Type.GetType() == typeof(Ois) && ((Ois)x.Type).ReferenceIndex.Name == index) ||
                    (x.Type.GetType() == typeof(FxBasisSwap) && 
                     (((FxBasisSwap)x.Type).BaseIndex.Name == index) &&
                     (((FxBasisSwap)x.Type).BaseIndex.Name == spreadIndex)))
                .ToList();
    }

    public static async Task<List<QuoteValue>> GetAllSwapCurveQuotes(
             string index,
             Tenor tenor,
             DateTime marketDataDate)
     {
         RateIndex underlyingRateIndex = new(index.Replace("_", "-"), tenor);
         OmicronClient client = new(@"https://omicron.fsa-aks.deloitte.co.za/");
 
         List<QuoteValue> curveQuotes = new();
         Dictionary<string, object> rateIndicesQuery = new() { ["$type"] = nameof(RateIndex) };
         
         QuoteType[] rateIndexTypes = await client.SearchAsync(rateIndicesQuery).ToArrayAsync();
         rateIndexTypes = 
             rateIndexTypes
                 .Where(x => 
                     ((RateIndex) x).Name.IgnoreCaseEquals(index.Replace("_", "-")) && 
                     ((RateIndex)x).Tenor.ToQuantLibPeriod() <= underlyingRateIndex.Tenor.ToQuantLibPeriod())
                 .ToArray();
 
         if (rateIndexTypes.Length != 0)
         {
             (QuoteType QuoteType, QuoteDto[] Quotes)[] rateIndexQuotes = 
                 await client.GetQuotesAsync(rateIndexTypes, marketDataDate, marketDataDate).ToArrayAsync(); 
             
             foreach ((QuoteType QuoteType, QuoteDto[] Quotes) rateIndexQuote in rateIndexQuotes)
             {
                 curveQuotes.Add(new(rateIndexQuote.QuoteType, rateIndexQuote.Quotes[0].Date, rateIndexQuote.Quotes[0].Quote));
             }
         }
         
         Dictionary<string, object> fraQuery = new() { ["$type"] = nameof(Fra) };
         QuoteType[] fraTypes = await client.SearchAsync(fraQuery).ToArrayAsync();
         fraTypes = fraTypes.Where(x => ((Fra) x).ReferenceIndex == underlyingRateIndex).ToArray();
 
         if (fraTypes.Length != 0)
         {
             (QuoteType QuoteType, QuoteDto[] Quotes)[] fraQuotes = 
                 await client.GetQuotesAsync(fraTypes, marketDataDate, marketDataDate).ToArrayAsync(); 
             
             foreach ((QuoteType QuoteType, QuoteDto[] Quotes) fraQuote in fraQuotes)
             {
                 curveQuotes.Add(new(fraQuote.QuoteType, fraQuote.Quotes[0].Date, fraQuote.Quotes[0].Quote));
             }
         }
 
         Dictionary<string, object> interestRateSwapQuery = new() { ["$type"] = nameof(InterestRateSwap) };
         QuoteType[] interestRateSwapTypes = await client.SearchAsync(interestRateSwapQuery).ToArrayAsync();
         interestRateSwapTypes = 
             interestRateSwapTypes.Where(x => ((InterestRateSwap)x).ReferenceIndex == underlyingRateIndex).ToArray();
         
         if (interestRateSwapTypes.Length != 0)
         {
             (QuoteType QuoteType, QuoteDto[] Quotes)[] interestRateSwapQuotes = 
                 await client.GetQuotesAsync(interestRateSwapTypes, marketDataDate, marketDataDate).ToArrayAsync();
             
             foreach ((QuoteType QuoteType, QuoteDto[] Quotes) interestRateSwapQuote in interestRateSwapQuotes)
             {
                 curveQuotes.Add(
                     new(
                         Type: interestRateSwapQuote.QuoteType, 
                         Date: interestRateSwapQuote.Quotes[0].Date, 
                         Value: interestRateSwapQuote.Quotes[0].Quote));
             }
         }
         
         Dictionary<string, object> oisQuery = new() { ["$type"] = nameof(Ois) };
         QuoteType[] oisTypes = await client.SearchAsync(oisQuery).ToArrayAsync();
         oisTypes = 
             oisTypes.Where(x => ((Ois)x).ReferenceIndex == underlyingRateIndex).ToArray();
         
         if (oisTypes.Length != 0)
         {
             (QuoteType QuoteType, QuoteDto[] Quotes)[] oisQuotes = 
                 await client.GetQuotesAsync(oisTypes, marketDataDate, marketDataDate).ToArrayAsync();
             
             foreach ((QuoteType QuoteType, QuoteDto[] Quotes) oisQuote in oisQuotes)
             {
                 curveQuotes.Add(
                     new QuoteValue(
                         Type: oisQuote.QuoteType, 
                         Date: oisQuote.Quotes[0].Date, 
                         Value: oisQuote.Quotes[0].Quote));
             }
         }
         
         return curveQuotes;
     }
    
    public static async Task<List<QuoteValue>> GetAllFxBasisCurveQuotes(
             string spreadIndexName,
             Tenor spreadIndexTenor,
             string baseIndexName,
             Tenor baseIndexTenor,
             Currency numeratorCurrency,
             Currency denominatorCurrency,
             DateTime marketDataDate)
    {
         RateIndex spreadIndex = new(spreadIndexName.Replace("_", "-"), spreadIndexTenor);
         RateIndex baseIndex = new(baseIndexName.Replace("_", "-"), baseIndexTenor);
         OmicronClient client = new(@"https://omicron.fsa-aks.deloitte.co.za/");

         FxSpot fxSpot = new(numeratorCurrency, denominatorCurrency);
         
         List<QuoteValue> curveQuotes = new();
         Dictionary<string, object> fxSpotQuery = new() { ["$type"] = nameof(FxSpot) };
         
         QuoteType[] fxSpotTypes = await client.SearchAsync(fxSpotQuery).ToArrayAsync();
         fxSpotTypes =
             fxSpotTypes
                 .Where(x => ((FxSpot) x) == fxSpot).ToArray();
 
         if (fxSpotTypes.Length != 0)
         {
             (QuoteType QuoteType, QuoteDto[] Quotes)[] fxSpotQuotes = 
                 await client.GetQuotesAsync(fxSpotTypes, marketDataDate, marketDataDate).ToArrayAsync(); 
             
             foreach ((QuoteType QuoteType, QuoteDto[] Quotes) fxSpotQuote in fxSpotQuotes)
             {
                 curveQuotes.Add(new(fxSpotQuote.QuoteType, fxSpotQuote.Quotes[0].Date, fxSpotQuote.Quotes[0].Quote));
             }
         }
         
         Dictionary<string, object> fxForwardQuery = new() { ["$type"] = nameof(FxForward) };
         QuoteType[] fxForwardTypes = await client.SearchAsync(fxForwardQuery).ToArrayAsync();
         fxForwardTypes = fxForwardTypes.Where(x => ((FxForward) x).FxSpot == fxSpot).ToArray();
         
         if (fxForwardTypes.Length != 0)
         {
             (QuoteType QuoteType, QuoteDto[] Quotes)[] fxForwardQuotes = 
                 await client.GetQuotesAsync(fxForwardTypes, marketDataDate, marketDataDate).ToArrayAsync(); 
             
             foreach ((QuoteType QuoteType, QuoteDto[] Quotes) fxForwardQuote in fxForwardQuotes)
             {
                 curveQuotes.Add(new(fxForwardQuote.QuoteType, fxForwardQuote.Quotes[0].Date, fxForwardQuote.Quotes[0].Quote));
             }
         }
 
         Dictionary<string, object> fxBasisSwapQuery = new() { ["$type"] = nameof(FxBasisSwap) };
         QuoteType[] fxBasisSwapTypes = await client.SearchAsync(fxBasisSwapQuery).ToArrayAsync();
         fxBasisSwapTypes = 
             fxBasisSwapTypes.Where(x => 
                 ((FxBasisSwap)x).BaseIndex == baseIndex && ((FxBasisSwap)x).SpreadIndex == spreadIndex).ToArray();
         
         if (fxBasisSwapTypes.Length != 0)
         {
             (QuoteType QuoteType, QuoteDto[] Quotes)[] fxBasisSwapQuotes = 
                 await client.GetQuotesAsync(fxBasisSwapTypes, marketDataDate, marketDataDate).ToArrayAsync();
             
             foreach ((QuoteType QuoteType, QuoteDto[] Quotes) fxBasisSwapQuote in fxBasisSwapQuotes)
             {
                 curveQuotes.Add(
                     new(
                         Type: fxBasisSwapQuote.QuoteType, 
                         Date: fxBasisSwapQuote.Quotes[0].Date, 
                         Value: fxBasisSwapQuote.Quotes[0].Quote));
             }
         }
         
         return curveQuotes;
    }

    /// <summary>
    /// Extracts the relevant FX vol surface quotes from a either Omicron directly or a pre-deserialized list of Omicron
    /// 'QuoteValues'. Thus one must either populate the <param name="quotes"/> or populate both the
    /// <param name="requisitionId"/> and the <param name="date"/>.
    /// </summary>
    /// <param name="quotes">(Optional)The list of Omicron quote values to loop through. If the</param>
    /// <param name="requisitionId">The relevant requisition ID in Omicron.</param>
    /// <returns>A list of the relevant swap curve quotes.</returns>
    public static List<QuoteValue> GetFxVolQuotes(
        List<QuoteValue>? quotes = null, 
        int? requisitionId = null,
        string? date = null)
    {
        quotes ??= GetOmicronRequisitionData(requisitionId, date);

        return 
            quotes
                .Where(x =>
                    (x.Type.GetType() == typeof(FxOption)) &&
                    ((FxOption)x.Type).ReferenceSpot == new FxSpot(Currency.USD, Currency.ZAR))
                .ToList();
    }

    /// <summary>
    /// Extracts a list of deserialized Omicron quote values from the Omicron REST APi for a given requisition ID and
    /// market data date.
    /// </summary>
    /// <param name="requisitionId">The relevant numerical requisition ID.</param>
    /// <param name="marketDataDate">Market data date.</param>
    /// <returns>A list of deserialized Omicron quote values.</returns>
    private static List<QuoteValue>? GetOmicronRequisitionData(int? requisitionId, string? marketDataDate)
    {
        if (requisitionId == null && marketDataDate == null)
        {
            return null;
        }

        using HttpClient client = new();
        client.DefaultRequestHeaders.Accept.Clear(); 
        client.BaseAddress = new Uri(OmicronUrl);
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));   
        HttpResponseMessage response = 
            client.GetAsync(OmicronUrl + "/api/requisition/" + requisitionId + "/" + marketDataDate).Result;
        
        if (response.IsSuccessStatusCode)
        {
            string jsonQuotes = response.Content.ReadAsStringAsync().Result;  
            return DeserializeOmicronObjects(jsonQuotes);
        }

        return null;
    }

    public abstract class JsonCreationConverter<T> : JsonConverter
    {
        protected abstract T Create(Type objectType, JObject jObject);
  
        public override bool CanConvert(Type objectType)
        {
            return typeof(T) == objectType;
        }
  
        public override object ReadJson(
            JsonReader reader, 
            Type objectType, 
            object? existingValue, 
            JsonSerializer serializer)
        {
            try
            {
                JObject jObject = JObject.Load(reader);
                T target = Create(objectType, jObject);
                serializer.Populate(jObject.CreateReader(), target);
                return target;
            }
            catch (JsonReaderException)
            {
                return null;
            }
        }
  
        public override void WriteJson(JsonWriter writer, object? value, JsonSerializer serializer)
        {
            throw new NotImplementedException();
        }
    }

    private class DExcelQuoteTypeConverter : JsonCreationConverter<QuoteType>
    {
        protected override QuoteType Create(Type objectType, JObject jObject)
        {
            Commodity commodity;
            FxSpot? fxSpot;
            int delta;
            RateIndex? rateIndex;
            Tenor? tenor;
            
            switch (jObject["$type"]?.ToString())
            {
                case "CommodityFuture":
                    tenor = JsonConvert.DeserializeObject<Tenor>(jObject["Tenor"]?.ToString() ?? string.Empty);
                    commodity = Enum.Parse<Commodity>(jObject["Commodity"]?.ToString() ?? string.Empty);
                    return new CommodityFuture(tenor, commodity);
                case "CommodityOption":
                    delta = jObject["Delta"].ToObject<int>();
                    Option option = Enum.Parse<Option>(jObject["OptionType"]?.ToString() ?? string.Empty); 
                    tenor = JsonConvert.DeserializeObject<Tenor>(jObject["Tenor"]?.ToString() ?? string.Empty);
                    commodity = Enum.Parse<Commodity>(jObject["Commodity"]?.ToString() ?? string.Empty);
                    return new CommodityOption(delta, option, tenor, commodity);
                case "Fra":
                    tenor = JsonConvert.DeserializeObject<Tenor>(jObject["Tenor"]?.ToString() ?? string.Empty);
                    rateIndex = JsonConvert.DeserializeObject<RateIndex>(jObject["ReferenceIndex"]?.ToString() ?? string.Empty);
                    return new Fra(tenor, rateIndex);
                case "FxBasisSwap":
                    rateIndex = JsonConvert.DeserializeObject<RateIndex>(jObject["BaseIndex"]?.ToString() ?? string.Empty);
                    RateIndex? spreadIndex = JsonConvert.DeserializeObject<RateIndex>(jObject["SpreadIndex"]?.ToString() ?? string.Empty);
                    tenor = JsonConvert.DeserializeObject<Tenor>(jObject["Tenor"]?.ToString() ?? string.Empty);
                    return new FxBasisSwap(rateIndex, spreadIndex, tenor);
                case "FxForward":
                    fxSpot = JsonConvert.DeserializeObject<FxSpot>(jObject["FxSpot"]?.ToString() ?? string.Empty);
                    tenor = JsonConvert.DeserializeObject<Tenor>(jObject["Tenor"]?.ToString() ?? string.Empty);
                    return new FxForward(fxSpot, tenor);
                case "FxOption":
                    delta = jObject["Delta"].ToObject<int>(); 
                    tenor = JsonConvert.DeserializeObject<Tenor>(jObject["Tenor"]?.ToString() ?? string.Empty);
                    fxSpot = JsonConvert.DeserializeObject<FxSpot>(jObject["ReferenceSpot"]?.ToString() ?? string.Empty);
                    return new FxOption(delta, tenor, fxSpot);    
                case "FxSpot":
                    Currency numerator = Enum.Parse<Currency>(jObject["Numerator"]?.ToString() ?? string.Empty);
                    Currency denominator = Enum.Parse<Currency>(jObject["Denominator"]?.ToString() ?? string.Empty);
                    return new FxSpot(numerator, denominator); 
                case "InterestRateSwap":
                    rateIndex = JsonConvert.DeserializeObject<RateIndex>(jObject["ReferenceIndex"]?.ToString() ?? string.Empty);
                    Tenor? paymentFrequency = JsonConvert.DeserializeObject<Tenor>(jObject["PaymentFrequency"]?.ToString() ?? string.Empty);
                    tenor = JsonConvert.DeserializeObject<Tenor>(jObject["Tenor"]?.ToString() ?? string.Empty);
                    return new InterestRateSwap(rateIndex, paymentFrequency, tenor);
                case "Ois":
                    rateIndex = JsonConvert.DeserializeObject<RateIndex>(jObject["ReferenceIndex"]?.ToString() ?? string.Empty);
                    tenor = JsonConvert.DeserializeObject<Tenor>(jObject["Tenor"]?.ToString());
                    return new Ois(rateIndex, tenor);
                case "RateIndex":
                    string? name = jObject["Name"]?.ToString();
                    tenor = JsonConvert.DeserializeObject<Tenor>(jObject["Tenor"]?.ToString() ?? string.Empty);
                    return new RateIndex(name, tenor);
            }
            
            return null;
        }
    }

    private class DExcelQuoteValueConverter : JsonCreationConverter<QuoteValue>
    {
        protected override QuoteValue Create(Type objectType, JObject jObject)
        {
            QuoteType? quoteType = JsonConvert.DeserializeObject<QuoteType>(jObject["type"]?.ToString() ?? string.Empty);
            DateTime date = DateTime.Parse(jObject["date"]?.ToString() ?? string.Empty);        
            double value = double.Parse(jObject["value"]?.ToString() ?? string.Empty);
            return new QuoteValue(quoteType, date, value); 
        }
    } 
}
