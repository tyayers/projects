using Microsoft.Azure;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace SupportSearch.Helper
{
    [Serializable]
    public class AzureFileUploadHelper : SupportSearch.Common.Interfaces.ISupportBlobUpload
    {
        protected string rootFilePath = CloudConfigurationManager.GetSetting("FilePath");

        public void UploadDocHtml(string Name, string HtmlContent, string KeyWords, string TextContent)
        {
            string htmlPath = rootFilePath + "\\html";
            if (!System.IO.Directory.Exists(htmlPath)) System.IO.Directory.CreateDirectory(htmlPath);
            System.IO.File.WriteAllText(htmlPath + "\\" + Name + ".html", HtmlContent);
        }

        public void UploadDocMarkdown(string Name, String MdContent, string KeyWords, string TextContent)
        {
            string mdPath = rootFilePath + "\\md";
            if (!System.IO.Directory.Exists(mdPath)) System.IO.Directory.CreateDirectory(mdPath);
            System.IO.File.WriteAllText(mdPath + "\\" + Name + ".md", MdContent);
        }

        public string UploadDocImage(string Name, Stream FileStream)
        {
            string imagePath = rootFilePath + "\\images";
            if (!System.IO.Directory.Exists(imagePath)) System.IO.Directory.CreateDirectory(imagePath);

            using (var fileStream = File.Create(imagePath + "\\" + Name + ".png"))
            {
                FileStream.Seek(0, SeekOrigin.Begin);
                FileStream.CopyTo(fileStream);
            }

            return "../images/" + Name + ".png";
        }

        public void UploadDocQnAMaker(string Name, string QnAContent)
        {
            string qnapath = rootFilePath + "\\qnamaker";
            if (!System.IO.Directory.Exists(qnapath)) System.IO.Directory.CreateDirectory(qnapath);

            System.IO.File.WriteAllText(qnapath + "\\" + Name + ".tsv", QnAContent);
        }
    }
}
