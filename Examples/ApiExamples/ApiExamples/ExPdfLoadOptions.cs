// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Loading;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    internal class ExPdfLoadOptions : ApiExampleBase
    {
        [TestCase(true)]
        [TestCase(false)]
        public void SkipPdfImages(bool isSkipPdfImages)
        {
            //ExStart
            //ExFor:PdfLoadOptions
            //ExFor:PdfLoadOptions.SkipPdfImages
            //ExFor:PdfLoadOptions.PageIndex
            //ExFor:PdfLoadOptions.PageCount
            //ExSummary:Shows how to skip images during loading PDF files.
            PdfLoadOptions options = new PdfLoadOptions();
            options.SkipPdfImages = isSkipPdfImages;
            options.PageIndex = 0;
            options.PageCount = 1;

            Document doc = new Document(MyDir + "Images.pdf", options);
            NodeCollection shapeCollection = doc.GetChildNodes(NodeType.Shape, true);

            if (isSkipPdfImages)
                Assert.AreEqual(shapeCollection.Count, 0);
            else
                Assert.AreNotEqual(shapeCollection.Count, 0);
            //ExEnd
        }
    }
}
