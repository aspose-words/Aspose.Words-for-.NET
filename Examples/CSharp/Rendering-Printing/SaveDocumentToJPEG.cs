using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class SaveDocumentToJPEG
    {
        public static void Run()
        {
            // ExStart:SaveDocumentToJPEG
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting();

            // Open the document
            Document doc = new Document(dataDir + "Rendering.doc");
            // Save as a JPEG image file with default options
            doc.Save(dataDir + "Rendering.JpegDefaultOptions.jpg");

            // Save document to stream as a JPEG with default options
            MemoryStream docStream = new MemoryStream();
            doc.Save(docStream, SaveFormat.Jpeg);
            // Rewind the stream position back to the beginning, ready for use
            docStream.Seek(0, SeekOrigin.Begin);

            // Save document to a JPEG image with specified options.
            // Render the third page only and set the JPEG quality to 80%
            // In this case we need to pass the desired SaveFormat to the ImageSaveOptions constructor 
            // to signal what type of image to save as.
            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Jpeg);
            imageOptions.PageIndex = 2;
            imageOptions.PageCount = 1;
            imageOptions.JpegQuality = 80;
            doc.Save(dataDir + "Rendering.JpegCustomOptions.jpg", imageOptions);
            // ExEnd:SaveDocumentToJPEG
        }
    }
}
