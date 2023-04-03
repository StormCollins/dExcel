namespace dExcel.OmicronUtils;

using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Linq;
using Omicron;
using JsonConverter = Newtonsoft.Json.JsonConverter;
using Option = Omicron.Option;

public static class OmicronUtils
{
    private const string OmicronUrl = "https://omicron.fsa-aks.deloitte.co.za";

    private const int RequisitionId = 1;

    private const string Date = "2023-02-14";

    public class OmicronObject
    {
        public QuoteType type;
        public DateTime date;
        public double value;

        public OmicronObject(QuoteType type, DateTime date, double value)
        {
            this.type = type;
            this.date = date;
            this.value = value;
        }
    }
    
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

    public static List<QuoteValue> GetSwapCurveQuotes(
        string indexName, 
        List<QuoteValue>? quotes = null, 
        int? requisitionId = null,
        string? date = null)
    {
        quotes ??= GetOmicronRequisitionData(requisitionId, date);

        return 
            quotes
                .Where(x => 
                    (x.Type.GetType() == typeof(RateIndex) && ((RateIndex) x.Type).Name == indexName) ||
                    (x.Type.GetType() == typeof(Fra) && ((Fra)x.Type).ReferenceIndex.Name == indexName) ||
                    (x.Type.GetType() == typeof(InterestRateSwap) && 
                     ((InterestRateSwap)x.Type).ReferenceIndex.Name == indexName) ||
                    (x.Type.GetType() == typeof(Ois) && ((Ois)x.Type).ReferenceIndex.Name == indexName))
                .ToList();
    }

    private static List<QuoteValue>? GetOmicronRequisitionData(int? requisitionId, string? date)
    {
        if (requisitionId == null && date == null)
        {
            return null;
        }

        using HttpClient client = new();
        client.DefaultRequestHeaders.Accept.Clear(); 
        client.BaseAddress = new Uri(OmicronUrl);
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));   
        HttpResponseMessage response = client.GetAsync(OmicronUrl + "/api/requisition/" + requisitionId + "/" + date).Result;
        
        if (response.IsSuccessStatusCode)
        {
            string jsonQuotes = response.Content.ReadAsStringAsync().Result;  //Make sure to add a reference to System.Net.Http.Formatting.dll
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
  
        public override object ReadJson(JsonReader reader, Type objectType,
            object existingValue, JsonSerializer serializer)
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
  
        public override void WriteJson(JsonWriter writer, object value,
            JsonSerializer serializer)
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
                    string empty = string.Empty;
                    if (empty != null)
                        tenor = JsonConvert.DeserializeObject<Tenor>(jObject["Tenor"]?.ToString() ?? empty);
                    return new Ois(new RateIndex("FEDFUND",new Tenor(1, TenorUnit.Day)), new Tenor(10, TenorUnit.Month));
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
