using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SupportSearch.Common.Dtos
{
    public class AzureSearchResult
    {
        [JsonProperty("@odata.context")]
        public string odatacontext { get; set; }
        public SearchDocument[] value { get; set; }
    }

    public class SearchDocument
    {
        [JsonProperty("@search.score")]
        public string searchscore { get; set; }
        public string searchkeywords { get; set; }
        public string metadata_storage_name { get; set; }
        public string metadata_storage_path { get; set; }
    }
}
