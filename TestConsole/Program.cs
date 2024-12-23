using RestSharp;
using System.Net;

namespace TestConsole
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // ServicePointManager.ServerCertificateValidationCallback = (sender, cert, chain, error) =>
            // {
            //     return true;
            // };

            // ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

            // ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;

            // ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

            // var client = new RestClient("http://hanwenxingming.com/api/translate/chinese-name");
            // var request = new RestRequest(Method.POST);
            // request.AddOrUpdateParameter("name", "衡阳");
            // var res = client.Execute(request);
            // if (res.StatusCode == HttpStatusCode.OK)
            // {
            //     Console.WriteLine(res.Content);
            // }
            // else
            // {
            //     Console.WriteLine(res.StatusCode);
            //     Console.WriteLine(res.ErrorMessage);
            //     Console.WriteLine(res.StatusDescription);
            // }
            // Console.WriteLine();

            HttpClientHandler handler = new HttpClientHandler();
            // Ignore SSL certificate errors
            handler.ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true;

            using (var client = new HttpClient(handler))
            {
                var url = "http://hanwenxingming.com/api/translate/chinese-name?name=衡阳";

                var values = new Dictionary<string, string>
                {
                    { "name", "衡阳" }
                };

                var content = new FormUrlEncodedContent(values);

                var response = client.PostAsync(url, content).Result;

                if (response.IsSuccessStatusCode)
                {
                    var responseString = response.Content.ReadAsStringAsync().Result;
                    Console.WriteLine(responseString);
                }
                else
                {
                    Console.WriteLine("Error: " + response.StatusCode);
                }
            }
        }
    }
}
