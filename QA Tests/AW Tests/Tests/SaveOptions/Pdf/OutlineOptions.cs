using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace QA_Tests.Tests.SaveOptions.Pdf
{
    /// <summary>
    /// Tests that verify saving to pdf using "CreateMissingOutlineLevels" parameter in "PdfSaveOptions"
    /// </summary>
    [TestFixture]
    internal class OutlineOptions : QaTestsBase
    {
        //Note: Test doesn't containt validation result, because it's difficult 
        //For validation result, you can save the document to pdf file and check out, that all bookmarks are created correctly for missing headings
        [Test]
        public void CreateMissingOutlineLevels()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            //Set maximum value of levels of headings
            pdfSaveOptions.OutlineOptions.HeadingsOutlineLevels = 9;
            pdfSaveOptions.OutlineOptions.CreateMissingOutlineLevels = true;
            pdfSaveOptions.SaveFormat = SaveFormat.Pdf;

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, pdfSaveOptions);

            dstStream.Dispose();
        }
    }
}
