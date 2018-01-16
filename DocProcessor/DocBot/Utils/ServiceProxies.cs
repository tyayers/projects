using DocBot.Dtos;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Xml.Linq;

namespace DocBot.Utils
{
    public static class ServiceProxies
    {
        public static async Task<QnAResponse> GetQnAResponse(string Question)
        {
            QnARequest request = new QnARequest();
            request.question = Question;

            QnAResponse data = new QnAResponse();

            using (WebClient client = new WebClient())
            {
                client.Headers.Add("Content-Type", "application/json");
                client.Headers.Add("Ocp-Apim-Subscription-Key", "6a50acd8314e456191ecf4cbf40164c9");
                client.Encoding = System.Text.Encoding.UTF8;

                string requestUri = "https://westus.api.cognitive.microsoft.com/qnamaker/v2.0/knowledgebases/efe61b54-60d9-4630-a773-2c9ae7288fed/generateAnswer";

                string responseString = client.UploadString(requestUri, JsonConvert.SerializeObject(request));
                string decodedString = System.Net.WebUtility.HtmlDecode(responseString).Replace("\"", "'").Replace("„", "'");
                data = JsonConvert.DeserializeObject<QnAResponse>(decodedString);
            }

            return data;
        }

        public static string SearchMarkdown(string search)
        {
            WebClient client = new WebClient();

            string response = client.DownloadString("https://botthedocs.search.windows.net/indexes/mdindex/docs?api-version=2016-09-01&search=" + search + "&api-key=CEB3B7D9BDA97F0434FD53986E9E2620");
            return response;
        }
    }
}