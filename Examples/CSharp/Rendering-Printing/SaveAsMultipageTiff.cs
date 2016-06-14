
using System;
using System.IO;
using System.Reflection;

using Aspose.Words;
using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class SaveAsMultipageTiff
    {
        public static void Run()
        {
            // ExStart:SaveAsMultipageTiff
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

            // Open the document.
            Document doc = new Document(dataDir + "TestFile Multipage TIFF.doc");

            //ExStart:SaveAsTIFF
            // Save the document as multipage TIFF.
            doc.Save(dataDir + "TestFile Multipage TIFF_out_.tiff");
            //ExEnd:SaveAsTIFF
            //ExStart:SaveAsTIFFUsingImageSaveOptions
            //Create an ImageSaveOptions object to pass to the Save method
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
            options.PageIndex = 0;
            options.PageCount = 2;
            options.TiffCompression = TiffCompression.Ccitt4;
            options.Resolution = 160;
            dataDir = dataDir + "TestFileWithOptions_out_.tiff";
            doc.Save(dataDir, options);
            //ExEnd:SaveAsTIFFUsingImageSaveOptions
            // ExEnd:SaveAsMultipageTiff
            Console.WriteLine("\nDocument saved as multi-page TIFF successfully.\nFile saved at " + dataDir);
        }
    }
}
