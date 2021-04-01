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
            //ExFor:PdfLoadOptions.SkipPdfImages
            //ExSummary:Shows how to skip images during loading PDF files.
            PdfLoadOptions options = new PdfLoadOptions();
            options.SkipPdfImages = isSkipPdfImages;
            
            Document doc = new Document(MyDir + "Images.pdf", options);
            NodeCollection shapeCollection = doc.GetChildNodes(NodeType.Shape, true);

            if (isSkipPdfImages)
            {
                Assert.AreEqual(shapeCollection.Count, 0);
            }
            else
            {
                Assert.AreNotEqual(shapeCollection.Count, 0);
            }
            //ExEnd
        }
    }
}
