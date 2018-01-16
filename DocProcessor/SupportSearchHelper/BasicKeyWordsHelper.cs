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
    public class BasicKeyWordsHelper : SupportSearch.Common.Interfaces.ISupportKeyWords
    {
        public string GetKeyWords(string Text)
        {
            string keyWords = "";

            keyWords = keyWords.ToLower().Replace(" the ", "").Replace(" a ", "");
     
            return keyWords;
        }
    }
}
