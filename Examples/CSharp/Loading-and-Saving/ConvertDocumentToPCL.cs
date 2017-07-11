using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class ConvertDocumentToPCL
    {
        public static void Run()
        {
            // ExStart:ConvertDocumentToPCL
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            // Load the document from disk.
            Document doc = new Document(dataDir + "Test File (docx).docx");

            PclSaveOptions saveOptions = new PclSaveOptions();

            saveOptions.SaveFormat = SaveFormat.Pcl;
            saveOptions.RasterizeTransformedElements = false;

            // Export the document as an PCL file.
            doc.Save(dataDir + "Document.PclConversion_out.pcl", saveOptions);
            // ExEnd:ConvertDocumentToPCL

            Console.WriteLine("\nDocument converted to PCL successfully.");
        }
    }
}
