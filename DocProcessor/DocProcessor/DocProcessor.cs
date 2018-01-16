using System.IO;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using System;
using SupportSearch.Helper;
using Autofac;
using SupportSearch.Common.Interfaces;

namespace DocProcessor
{
    public static class DocProcessorFunction
    {
        // DI container
        private static IContainer Container { get; set; }

        // Constructor
        static DocProcessorFunction()
        {
            var builder = new ContainerBuilder();

            // Register individual components
            builder.RegisterInstance(new AzureFileUploadHelper())
                   .As<ISupportBlobUpload>();
            builder.RegisterInstance(new AzureMLKeyWordsHelper())
                   .As<ISupportKeyWords>();
            builder.RegisterType<WordConverter>();
            Container = builder.Build();
        }

        [FunctionName("DocFunction")]
        public static void Run([BlobTrigger("docs/{name}", Connection = "SupportDocsConnection")]Stream myBlob, string name, TraceWriter log)
        {
            log.Info($"C# Blob trigger function Processed blob\n Name: \n Size: {myBlob.Length} Bytes");

            System.DateTime startTime = DateTime.Now;
            SupportSearch.Helper.WordConverter converter = Container.Resolve<WordConverter>();
            converter.ConvertDocToHtml(name, myBlob);

            log.Info($"Finished processing word file {name} in {System.DateTime.Now.Subtract(startTime).ToString()}.");
        }
    }
}