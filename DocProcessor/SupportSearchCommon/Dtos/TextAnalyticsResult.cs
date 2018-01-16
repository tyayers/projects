using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SupportSearch.Common.Dtos
{
    public class TextAnalyticsResult
    {
        public TextAnalyticsResultDocument[] documents { get; set; }
    }

    public class TextAnalyticsResultDocument
    {
        public string[] keyPhrases { get; set; }
        public TextAnalyticsError[] errors { get; set; }
    }

    public class TextAnalyticsError
    {
        public string id { get; set; }
        public string message { get; set; }
    }
}
