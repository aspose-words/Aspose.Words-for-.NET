using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Loading_and_Saving
{
    class ConvertWordDocument
    {
        public static void Run()
        {
            ConvertDocumentToPNG();
        }

        // ExStart:ConvertDocumentToPNG
        public static void ConvertDocumentToPNG()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            // Load a document
            Document doc = new Document(dataDir + "SampleImages.doc");

            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Png);

            PageRange pageRange = new PageRange(0, doc.PageCount - 1);
            imageSaveOptions.PageSet = new PageSet(pageRange);
            imageSaveOptions.PageSavingCallback = new HandlePageSavingCallback();
            doc.Save(dataDir + "output.png", imageSaveOptions);
            
            Console.WriteLine("\nDocument converted to PNG successfully.");
        }

        private class HandlePageSavingCallback : IPageSavingCallback
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            public void PageSaving(PageSavingArgs args)
            {
                args.PageFileName = string.Format(dataDir + "Page_{0}.png", args.PageIndex);
            }
        }
        // ExEnd:ConvertDocumentToPNG
    }
}
