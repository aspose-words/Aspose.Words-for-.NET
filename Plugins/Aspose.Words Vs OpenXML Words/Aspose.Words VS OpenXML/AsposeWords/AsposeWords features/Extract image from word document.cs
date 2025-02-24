// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

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
        public void ExtractImage()
        {
            Document doc = new Document(MyDir + "Extract image.docx");

            // Get the collection of shapes from the document,
            // and save the image data of every shape with an image as a file to the local file system.
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            int imageIndex = 0;
            foreach (Shape shape in shapes.OfType<Shape>())
            {
                if (shape.HasImage)
                {
                    // The image data of shapes may contain images of many possible image formats. 
                    // We can determine a file extension for each image automatically, based on its format.
                    string imageFileName =
                        $"ExtractImage.{imageIndex}.Aspose.Words{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}";
                    shape.ImageData.Save(ArtifactsDir + imageFileName);
                    imageIndex++;
                }
            }
        }
    }
}
