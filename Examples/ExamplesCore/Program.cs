using Aspose.Words;
using Aspose.Words.Saving;
using System;
using System.IO;

namespace ExamplesCore
{
    class Program
    {
        static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = GetDataDir_LoadingAndSaving();

            LoadPDF(dataDir);
            LoadAndSaveEncryptedPDF(dataDir);
            LoadPageRangeOfPDF(dataDir);
        }

        public static void LoadPDF(string dataDir)
        {
            //ExStart:LoadPDF
            Document doc = new Document(dataDir + "Document.pdf");

            dataDir = dataDir + "Document_out.pdf";
            doc.Save(dataDir);
            //ExEnd:LoadPDF
            Console.WriteLine("\nDocument saved.\nFile saved at " + dataDir);
        }

        public static void LoadAndSaveEncryptedPDF(string dataDir)
        {
            // ExStart:LoadAndSaveEncryptedPDF  
            PdfLoadOptions pdfLoad = new PdfLoadOptions();
            pdfLoad.Password = "password";

            Document doc = new Document(dataDir + @"encrypted.pdf", pdfLoad);
            doc.Save(dataDir + "out.pdf");
            // ExEnd:LoadAndSaveEncryptedPDF 
            Console.WriteLine("\nLoad and save encrypted document successfully.\nFile saved at " + dataDir);
        }

        public static void LoadPageRangeOfPDF(string dataDir)
        {
            // ExStart:LoadPageRangeOfPDF  
            PdfLoadOptions pdfLoad = new PdfLoadOptions();
            pdfLoad.PageIndex = 0;
            pdfLoad.PageCount = 2;

            Document doc = new Document(dataDir + @"Document1.pdf", pdfLoad);
            doc.Save(dataDir + "out.pdf");
            // ExEnd:LoadPageRangeOfPDF 
            Console.WriteLine("\nLoad and save PDF with specific Page Range successfully.\nFile saved at " + dataDir);
        }

        public static String GetDataDir_LoadingAndSaving()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Loading-and-Saving/");
        }
        private static string GetDataDir_Data()
        {
            var parent = Directory.GetParent(Directory.GetCurrentDirectory()).Parent;
            string startDirectory = null;
            if (parent != null)
            {
                var directoryInfo = parent.Parent;
                if (directoryInfo != null)
                {
                    var directoryPInfo = directoryInfo.Parent;
                    startDirectory = directoryPInfo.FullName;
                }
            }
            else
            {
                startDirectory = parent.FullName;
            }

            return Path.Combine(startDirectory, "Data\\");
        }


    }
}
