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
    public class AzureBlobUploadHelper : SupportSearch.Common.Interfaces.ISupportBlobUpload
    {
        public void UploadDocHtml(string Name, string HtmlContent, string KeyWords, string TextContent)
        {
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(CloudConfigurationManager.GetSetting("StorageConnectionString"));
            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();

            string containerName = CloudConfigurationManager.GetSetting("StorageHtmlContainer");
            // Retrieve a reference to a container.
            CloudBlobContainer container = blobClient.GetContainerReference(containerName);

            // Retrieve reference to the blob
            CloudBlockBlob blockBlob = container.GetBlockBlobReference(Name);

            blockBlob.UploadText(HtmlContent);

            KeyWords = KeyWords.Replace("+", " ").Replace("|", "");
            blockBlob.Metadata["searchkeywords"] = HttpServerUtility.UrlTokenEncode(Encoding.UTF8.GetBytes(KeyWords));
            blockBlob.Metadata["textcontent"] = HttpServerUtility.UrlTokenEncode(Encoding.UTF8.GetBytes(TextContent));

            try
            {
                blockBlob.SetMetadata();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public void UploadDocMarkdown(string Name, String MdContent, string KeyWords, string TextContent)
        {
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(CloudConfigurationManager.GetSetting("StorageConnectionString"));
            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();

            string containerName = CloudConfigurationManager.GetSetting("StorageMdContainer");

            // Retrieve a reference to a container.
            CloudBlobContainer container = blobClient.GetContainerReference(containerName);

            // Retrieve reference to a blob named "myblob".
            CloudBlockBlob blockBlob = container.GetBlockBlobReference(Name);

            blockBlob.UploadText(MdContent);

            KeyWords = KeyWords.Replace("+", " ").Replace("|", "");
            blockBlob.Metadata["searchkeywords"] = HttpServerUtility.UrlTokenEncode(Encoding.UTF8.GetBytes(KeyWords));
            blockBlob.Metadata["textcontent"] = HttpServerUtility.UrlTokenEncode(Encoding.UTF8.GetBytes(TextContent));

            try
            {
                blockBlob.SetMetadata();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public string UploadDocImage(string Name, Stream FileStream)
        {
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(CloudConfigurationManager.GetSetting("StorageConnectionString"));
            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();

            string containerName = CloudConfigurationManager.GetSetting("StorageImageContainer");

            // Retrieve a reference to a container.
            CloudBlobContainer container = blobClient.GetContainerReference(containerName);

            // Retrieve reference to a blob named "myblob".
            CloudBlockBlob blockBlob = container.GetBlockBlobReference(Name);

            blockBlob.UploadFromStream(FileStream);

            return blockBlob.Uri.ToString();
        }

        public void UploadDocQnAMaker(string Name, string QnAContent)
        {
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(CloudConfigurationManager.GetSetting("StorageConnectionString"));
            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();

            string containerName = CloudConfigurationManager.GetSetting("QnaMdContainer");

            // Retrieve a reference to a container.
            CloudBlobContainer container = blobClient.GetContainerReference(containerName);

            // Retrieve reference to a blob named "myblob".
            CloudBlockBlob blockBlob = container.GetBlockBlobReference(Name + ".tsv");

            blockBlob.UploadText(QnAContent);
        }
    }
}
