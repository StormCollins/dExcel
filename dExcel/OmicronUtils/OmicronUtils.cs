namespace dExcel.OmicronUtils;

using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json.Linq;
using Omicron;
using Omicron.Data.Serialisation;
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
    
    public static void SerializeOmicronObject()
    {
        OmicronObject omicronObject =
            new(
                type: new CommodityOption(25, Option.Put, new Tenor(12, TenorUnit.Month), Commodity.BrentCrudeIce),
                date: new DateTime(2023, 02, 14),
                value: 12);
        
        string json = JsonConvert.SerializeObject(omicronObject);

        string serializedObject =
            "[" +
            "{\"type\":{\"$type\":\"CommodityFuture\",\"Tenor\": {\"amount\":12,\"unit\":\"Month\"},\"Commodity\":\"Ethane\"},\"date\":\"2023-02-14T00:00:00\",\"value\":0.24625}," +
            "{\"type\":{\"$type\":\"CommodityOption\",\"Delta\":25,\"OptionType\":\"Put\",\"Tenor\":{\"amount\":10,\"unit\":\"Month\"},\"Commodity\":\"BrentCrudeIce\"},\"date\":\"2023-02-14T00:00:00\",\"value\":0.4103 }]";

        string commodityFuture =
            "{\"type\":{\"$type\":\"CommodityFuture\",\"Tenor\": {\"amount\":12,\"unit\":\"Month\"},\"Commodity\":\"Ethane\"},\"date\":\"2023-02-14T00:00:00\",\"value\":0.24625}";

        char x = commodityFuture[45];

        JsonConvert.DefaultSettings = () => new JsonSerializerSettings
        {
            Converters = new List<JsonConverter>
            {
                new dExcelQuoteTypeConverter(), 
                new StringEnumConverter(),
                new dExcelQuoteValueConverter(),
            }
        };
        
        QuoteTypeConverter converter = new QuoteTypeConverter();
        
        List<QuoteValue> deserializedObject = JsonConvert.DeserializeObject<List<QuoteValue>>(serializedObject);
         
    }


    public static List<QuoteValue> DeserializeOmicronObject(string json)
    {
        JsonConvert.DefaultSettings = () => new JsonSerializerSettings
        {
            Converters = new List<JsonConverter>
            {
                new dExcelQuoteTypeConverter(), 
                new StringEnumConverter(),
                new dExcelQuoteValueConverter(),
            }
        };
        
        QuoteTypeConverter converter = new QuoteTypeConverter();
        
        return JsonConvert.DeserializeObject<List<QuoteValue>>(json);
    }
    
    public static void PullData()
    {
        using HttpClient client = new HttpClient();
        client.DefaultRequestHeaders.Accept.Clear(); 
        client.BaseAddress = new Uri(OmicronUrl);
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));   
        HttpResponseMessage response = client.GetAsync(OmicronUrl + "/api/requisition/" + RequisitionId + "/" + Date).Result;
        if (response.IsSuccessStatusCode)
        {
            var dataObjects = response.Content.ReadAsStringAsync().Result;  //Make sure to add a reference to System.Net.Http.Formatting.dll
            // JsonConverter converter;
            // JsonConverterFactory jsonConverterFactory = 
        }
        else
        {
            Console.WriteLine("{0} ({1})", (int)response.StatusCode, response.ReasonPhrase);
        }
    }

    // public class OmicronObjectConverter<T> : JsonConverter<T> 
    //     where T : OmicronObject
    // {
    //     public override void WriteJson(JsonWriter writer, T? value, JsonSerializer serializer)
    //     {
    //         throw new NotImplementedException();
    //     }
    //
    //     public override T? ReadJson(JsonReader reader, Type objectType, T? existingValue, bool hasExistingValue, JsonSerializer serializer)
    //     {
    //         reader.Read();
    //         return null;
    //     }
    //     
    //     public override bool CanConvert(Type objectType)
    //     {
    //         return typeof(objectType).IsAssignableFrom(typeof(T));
    //     }
    // }
    
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
                var jObject = JObject.Load(reader);
                var target = Create(objectType, jObject);
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
    
    public class dExcelQuoteTypeConverter : JsonCreationConverter<QuoteType>
    {
        protected override QuoteType Create(Type objectType, JObject jObject)
        {
            Tenor tenor;
            Commodity commodity;
            int delta;
            switch (jObject["$type"].ToString())
            {
                case "CommodityFuture":
                    tenor = JsonConvert.DeserializeObject<Tenor>(jObject["Tenor"].ToString());
                    commodity = Enum.Parse<Commodity>(jObject["Commodity"].ToString());
                    return new CommodityFuture(tenor, commodity);
                case "CommodityOption":
                    delta = jObject["Delta"].ToObject<int>();
                    Option option = Enum.Parse<Option>(jObject["OptionType"].ToString()); 
                    tenor = JsonConvert.DeserializeObject<Tenor>(jObject["Tenor"].ToString());
                    commodity = Enum.Parse<Commodity>(jObject["Commodity"].ToString());
                    return new CommodityOption(delta, option, tenor, commodity);
                case "FxOption":
                    delta = jObject["Delta"].ToObject<int>(); 
                    tenor = JsonConvert.DeserializeObject<Tenor>(jObject["Tenor"].ToString());
                    FxSpot fxSpot = JsonConvert.DeserializeObject<FxSpot>(jObject["ReferenceSpot"].ToString());
                    return new FxOption(delta, tenor, fxSpot);    
                case "FxSpot":
                    Currency numerator = Enum.Parse<Currency>(jObject["Numerator"].ToString());
                    Currency denominator = Enum.Parse<Currency>(jObject["Denominator"].ToString());
                    return new FxSpot(numerator, denominator); 
                case "RateIndex":
                    string name = jObject["Name"].ToString();
                    tenor = JsonConvert.DeserializeObject<Tenor>(jObject["Tenor"].ToString());
                    return new RateIndex(name, tenor);
                case "InterestRateSwap":
                    RateIndex rateIndex = JsonConvert.DeserializeObject<RateIndex>(jObject["ReferenceIndex"].ToString());
                    Tenor paymentFrequency = JsonConvert.DeserializeObject<Tenor>(jObject["PaymentFrequency"].ToString());
                    tenor = JsonConvert.DeserializeObject<Tenor>(jObject["Tenor"].ToString());
                    return new InterestRateSwap(rateIndex, paymentFrequency, tenor);
            }
            return null;
        }
    } 
    
    
    public class dExcelQuoteValueConverter : JsonCreationConverter<QuoteValue>
    {
        protected override QuoteValue Create(Type objectType, JObject jObject)
        {
            QuoteType quoteType = JsonConvert.DeserializeObject<QuoteType>(jObject["type"].ToString());
            DateTime date = DateTime.Parse(jObject["date"].ToString());        
            double value = double.Parse(jObject["value"].ToString());
            return new QuoteValue(quoteType, date, value); 
        }
    } 
    
}
