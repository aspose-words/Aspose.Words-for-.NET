// Copyright (c) Aspose 2002-2021. All Rights Reserved.

/*
    This project uses NuGet's Automatic Package Restore feature to 
    resolve the Aspose.Words for .NET API reference when the project is built. 
    Please visit https://docs.nuget.org/consume/nuget-faq for more information. 

    If you do not wish to use NuGet, you can manually download Aspose.Words for .NET API 
    from http://www.aspose.com/downloads, install it, and then add a reference to it to this project. 

    For any issues, questions or suggestions, please visit the Aspose Forums: https://forum.aspose.com/
*/

using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.IO;
using System.Linq;

namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string File = FilePath + "Extract Image - Aspose.docx";

            Document doc = new Document(File);

            // Save the document to memory and reload it.
            using (MemoryStream stream = new MemoryStream())
            {
                doc.Save(stream, SaveFormat.Doc);
                Document doc2 = new Document(stream);

                // "Shape" nodes that have the "HasImage" flag set contain and display images.
                IEnumerable<Shape> shapes = doc2.GetChildNodes(NodeType.Shape, true)
                    .OfType<Shape>().Where(s => s.HasImage);

                int imageIndex = 0;
                foreach (Shape shape in shapes)
                {
                    string imageFileName = string.Format(
                        "Image.ExportImages.{0}_out_{1}", imageIndex,
                        FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType));

                    shape.ImageData.Save(FilePath + imageFileName);
                    imageIndex++;
                }
            }
        }
    }
}
