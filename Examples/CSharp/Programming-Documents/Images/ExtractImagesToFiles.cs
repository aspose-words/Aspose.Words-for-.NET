using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;
using Aspose.Words.Drawing;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Images
{
    class ExtractImagesToFiles
    {
        public static void Run()
        {
            //ExStart:ExtractImagesToFiles
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithImages();
            Document doc = new Document(dataDir + "Image.SampleImages.doc");

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = string.Format(
                        "Image.ExportImages.{0}_out_{1}", imageIndex, FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType));
                    shape.ImageData.Save(dataDir + imageFileName);
                    imageIndex++;
                }
            }
            //ExEnd:ExtractImagesToFiles
            Console.WriteLine("\nAll images extracted from document.");
        }
    }
}
