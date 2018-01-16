using SupportSearch.Common;
using SupportSearch.Common.Dtos;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace SupportSearch.Helper
{
    [Serializable]
    public class AzureCSKeyWordHandler : SupportSearch.Common.Interfaces.ISupportKeyWords
    {
        public string GetKeyWords(string Text)
        {
            List<string> results = new List<string>();

            try
            {
                string uri = "https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/keyPhrases";
                HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(uri);
                httpWebRequest.Method = WebRequestMethods.Http.Post;
                string subKey = System.Configuration.ConfigurationManager.AppSettings["TextAnalyticsKey"];

                //string postData = "{ documents: [ {language: 'de', id: '1', text: '" + HttpUtility.UrlEncode(Text, Encoding.UTF8) + "' } ] }";
                string postData = "{ documents: [ {language: 'de', id: '1', text: '" + Text + "' } ] }";
                var data = Encoding.UTF8.GetBytes(postData);

                httpWebRequest.Headers.Add("Ocp-Apim-Subscription-Key", subKey);
                using (var stream = httpWebRequest.GetRequestStream())
                {
                    stream.Write(data, 0, data.Length);
                }

                var response = (HttpWebResponse)httpWebRequest.GetResponse();

                var responseString = new StreamReader(response.GetResponseStream()).ReadToEnd();

                TextAnalyticsResult result = Newtonsoft.Json.JsonConvert.DeserializeObject<TextAnalyticsResult>(responseString);

                if (result.documents != null)
                {
                    foreach (TextAnalyticsResultDocument doc in result.documents)
                    {
                        foreach (string docText in doc.keyPhrases)
                        {
                            if (docText.Length > 3) results.Add(docText);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                results.AddRange(Text.Split(' '));
            }


            return String.Join(" ", results.ToArray<string>());
        }
    }
}
