using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace SupportSearch.Helper
{
    [Serializable]
    public class AzureMLKeyWordsHelper : SupportSearch.Common.Interfaces.ISupportKeyWords
    {
        public string GetKeyWords(string Text)
        {
            string keyWords = "";

            using (var client = new System.Net.Http.HttpClient())
            {
                var scoreRequest = new
                {
                    Inputs = new Dictionary<string, List<Dictionary<string, string>>>() {
                        {
                            "input1",
                            new List<Dictionary<string, string>>(){new Dictionary<string, string>(){
                                            {
                                                "Chapter", ""
                                            },
                                            {
                                                "Headers", Text
                                            },
                                }
                            }
                        },
                    },
                    GlobalParameters = new Dictionary<string, string>()
                    {
                    }
                };

                string apiKey = System.Configuration.ConfigurationManager.AppSettings["AzureMLSupportTextKey"];
                client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);
                client.BaseAddress = new Uri("https://europewest.services.azureml.net/subscriptions/fb3af244eaab4a8387d4fb10ce8a6260/services/9c9c434e9da64585b7c81a1b6bbe4905/execute?api-version=2.0&format=swagger");

                Task<HttpResponseMessage> responseTask = client.PostAsJsonAsync("", scoreRequest);
                responseTask.Wait();
                HttpResponseMessage response = responseTask.Result;

                if (response.IsSuccessStatusCode)
                {
                    Task<string> resultTask = response.Content.ReadAsStringAsync();
                    resultTask.Wait();
                    string result = resultTask.Result;

                    //keyWords = result;
                    JObject jsonResult = JObject.Parse(result);
                    keyWords = jsonResult["Results"]["output1"][0]["Preprocessed Headers"].Value<string>();
                    keyWords = keyWords.Replace("|", "");
                }
                else
                {
                    Console.WriteLine(string.Format("The Azure ML Key Word request failed with status code: {0}", response.StatusCode));
                    System.Diagnostics.Trace.TraceError(string.Format("The Azure ML Key Word request failed with status code: {0}", response.StatusCode));
                    // Print the headers - they include the requert ID and the timestamp,
                    // which are useful for debugging the failure
                    Console.WriteLine(response.Headers.ToString());
                }
            }

            // TODO parse json structure from machine learning
            //JObject jsonResult = JObject.Parse(keyWords);
            //docPart.KeyWords = jsonResult["Results"]["output1"][0]["Preprocessed Headers"].Value<string>().Replace("|", "");


            return keyWords;
        }
    }
}
