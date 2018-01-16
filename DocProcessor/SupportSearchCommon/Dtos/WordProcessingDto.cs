using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace SupportSearch.Common.Dtos
{
    public enum HeaderType{
        Header1,
        Header2,
        Header3,
        Header4,
        Header5,
        SubHeader
    };

    public class WordProcessingDocDto
    {
        public WordProcessingDocDto(string Name)
        {
            DocumentName = Name;
        }
        public string DocumentName { get; set; }
        public List<string> Heading1StyleIds { get; set; } = new List<string>();
        public List<string> Heading2StyleIds { get; set; } = new List<string>();
        public List<string> Heading3StyleIds { get; set; } = new List<string>();
        public List<string> Heading4StyleIds { get; set; } = new List<string>();
        public List<string> Heading5StyleIds { get; set; } = new List<string>();
        public List<WordProcessingDocPart> WordProcessingDocParts { get; set; } = new List<WordProcessingDocPart>();
    }

    public class WordProcessingDocPart
    {
        public WordProcessingDocPart(HeaderType type)
        {
            Type = type;
        }
        public HeaderType Type { get; set; }
        public string HeaderText { get; set; }
        public string HeaderNumber { get; set; }
        public string KeyWords { get; set; }
        public List<string> Content { get; set; } = new List<string>();
        public XElement XmlElement { get; set; }
        public string HtmlOutput { get; set; }
        public string MarkdownOutput { get; set; }
        public string TextOutput { get; set; }
        public string QnaOutput { get; set; }
        public WordProcessingDocPart ParentHeader1 { get; set; }
        public WordProcessingDocPart ParentHeader2 { get; set; }
        public WordProcessingDocPart ParentHeader3 { get; set; }
        public WordProcessingDocPart ParentHeader4 { get; set; }
        public WordProcessingDocPart ParentHeader5 { get; set; }
        public List<string> ImagePaths { get; set; } = new List<string>();
        public List<WordProcessingDocPart> SubHeaders { get; set; } = new List<WordProcessingDocPart>();
    }
}
