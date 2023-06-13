using System.Net.Http;
using System.Net.Http.Headers;
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
    public static List<QuoteValue> GetSwapCurveQuotes(
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
