using Autofac;
using SupportSearch.Common.Interfaces;
using SupportSearch.Helper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocTester
{
    class Program
    {
        private static IContainer Container { get; set; }

        static void Main(string[] args)
        {
            Init();

            //string docPath = "C:\\git\\BotTheDocsProject\\DocProcessor\\MicrosotFlowBasics.docx";
            string docPath = "C:\\git\\BotTheDocsProject\\DocProcessor\\MicrosoftFlowDocs.docx";

            if (System.IO.File.Exists(docPath))
            {
                try
                {
                    Console.WriteLine($"Converting document {docPath}");
                    System.IO.Stream strm = System.IO.File.OpenRead(docPath);
                    SupportSearch.Helper.WordConverter converter = Container.Resolve<WordConverter>();
                    converter.ConvertDocToHtml("HomeControlUserManual", strm);
                    Console.WriteLine("Conversion finished! Press any key to quit..");
                    Console.ReadLine();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error converting document:");
                    Console.WriteLine(ex.ToString());
                    Console.ReadLine();
                }
            }
            else
            {
                Console.WriteLine("Input document does not exist!");
                Console.ReadLine();
            }
        }

        private static void Init()
        {
            var builder = new ContainerBuilder();

            // Register individual components
            builder.RegisterInstance(new AzureFileUploadHelper())
                   .As<ISupportBlobUpload>();
            builder.RegisterInstance(new BasicKeyWordsHelper())
                   .As<ISupportKeyWords>();
            builder.RegisterType<WordConverter>();
            Container = builder.Build();
        }
    }
}
