using Microsoft.Extensions.Hosting;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace D365WebAPIWithPolly
{
    public class D365WebAPIService : IHostedService
    {
        private IHttpClientFactory _httpClientFactory;

        public D365WebAPIService(IHttpClientFactory httpClientFactory)
        {
            _httpClientFactory = httpClientFactory;
        }

        public async Task StartAsync(CancellationToken cancellationToken)
        {
            await WhoAmIRequestAsync();
        }

        public async Task WhoAmIRequestAsync()
        {
            HttpClient httpClient = _httpClientFactory.CreateClient("D365WebApiClient");
            var response = await httpClient.GetAsync("WhoAmI");

            if (response.IsSuccessStatusCode)
            {
                string json = await response.Content.ReadAsStringAsync();
                JObject result = JsonConvert.DeserializeObject(json) as JObject;
                Console.ForegroundColor = ConsoleColor.Gray;
                Display(result);
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Failed to call the Web Api: {response.StatusCode}");
                string content = await response.Content.ReadAsStringAsync();

                // Note that if you got reponse.Code == 403 and reponse.content.code == "Authorization_RequestDenied"
                // this is because the tenant admin as not granted consent for the application to call the Web API
                Console.WriteLine($"Content: {content}");
            }
        }
        private static void Display(JObject result)
        {
            foreach (JProperty child in result.Properties().Where(p => !p.Name.StartsWith("@")))
            {
                Console.WriteLine($"{child.Name} = {child.Value}");
            }
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            throw new NotImplementedException();
        }
    }
}
