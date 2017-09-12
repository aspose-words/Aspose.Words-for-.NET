using Aspose.Words.Rendering;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Rendering_Printing
{
    public class Print_CachePrinterSettings
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting();

            CachePrinterSettings(dataDir);
        }

        public static void CachePrinterSettings(string dataDir)
        {
            // ExStart:CachePrinterSettings  
            //Load the Word document
            Document doc = new Document(dataDir + "TestFile.doc");

            // Build layout.
            doc.UpdatePageLayout();

            // Create settings, setup printing.
            PrinterSettings settings = new PrinterSettings();
            settings.PrinterName = "Microsoft XPS Document Writer";

            // Create AsposeWordsPrintDocument  and cache settings.
            AsposeWordsPrintDocument printDocument = new AsposeWordsPrintDocument(doc);
            printDocument.PrinterSettings = settings;
            printDocument.CachePrinterSettings();

            printDocument.Print();

            // ExEnd:CachePrinterSettings
            Console.WriteLine("\nDocument is printed successfully.");
        }
    }
}
