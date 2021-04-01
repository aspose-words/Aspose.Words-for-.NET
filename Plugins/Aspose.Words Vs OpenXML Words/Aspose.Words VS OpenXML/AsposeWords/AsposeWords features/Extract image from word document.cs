// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class ExtractImageFromWordDocument : TestUtil
    {
        [Test]
        public void ExtractImageFromWordDocumentFeature()
        {
            Document doc = new Document(MyDir + "Extract image.docx");

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
                    string imageFileName =
                        $"Image.ExportImages.{imageIndex}_Aspose.Words_{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}";

                    shape.ImageData.Save(ArtifactsDir + imageFileName);
                    imageIndex++;
                }
            }
        }
    }
}
