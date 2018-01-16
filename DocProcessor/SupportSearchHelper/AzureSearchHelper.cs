using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SupportSearch.Common.Dtos;
using System.Net;

namespace SupportSearch.Helper
{
    [Serializable]
    public class AzureSearchHelper : SupportSearch.Common.Interfaces.ISupportSearch
    {
        public string SearchHtmlDocs(string search)
        {
            throw new NotImplementedException();
        }

        public string SearchMarkdownDocs(string search)
        {
            WebClient client = new WebClient();

            string response = client.DownloadString("https://botthedocs.search.windows.net/indexes/md-index/docs?api-version=2016-09-01&search=" + search + "&api-key=CEB3B7D9BDA97F0434FD53986E9E2620");
            return response;
        }

            //public AzureSearchResult SearchHtmlDocs(string search)
            //{
            //    throw new NotImplementedException();
            //}

            //public AzureSearchResult SearchMarkdownDocs(string search)
            //{
            //    WebClient client = new WebClient();

            //    string response = client.DownloadString("https://docs.search.windows.net/indexes/pageindex/docs?api-version=2016-09-01&search=" + search + "&api-key=007E8A611B46D0E06225E900A3594206");
            //    AzureSearchResult result = Newtonsoft.Json.JsonConvert.DeserializeObject<AzureSearchResult>(response);

            //    return result;
            //}
        }
    }
