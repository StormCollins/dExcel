namespace dExcel.OmicronUtils;

using System.Collections;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json.Serialization;
using SkiaSharp.HarfBuzz;
using Newtonsoft.Json;

public static class OmicronCurveUtils
{
    private const string OmicronUrl = "https://omicron.fsa-aks.deloitte.co.za";

    private const int RequisitionId = 1;

    private const string Date = "2023-02-14";

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
}
