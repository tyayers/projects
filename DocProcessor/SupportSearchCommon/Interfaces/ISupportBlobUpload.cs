using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SupportSearch.Common.Interfaces
{
    public interface ISupportBlobUpload
    {
        void UploadDocHtml(string Name, string HtmlContent, string KeyWords, string TextContent);

        void UploadDocMarkdown(string Name, string MdContent, string KeyWords, string TextContent);

        void UploadDocQnAMaker(string Name, string QnAContent);

        string UploadDocImage(string Name, Stream ImageStream);
    }
}
