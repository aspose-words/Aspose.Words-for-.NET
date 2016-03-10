using Aspose.Pdf.Facades;
using Aspose.Pdf.Text;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace QaTests.Tests
{
    /// <summary>
    /// Tests that verify saving to pdf with special save options
    /// </summary>
    [TestFixture]
    internal class QaPdfSaveOptions : QaTestsBase
    {
        [Test]
        public void CreateMissingOutlineLevels()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            //Set maximum value of levels of headings
            pdfSaveOptions.OutlineOptions.HeadingsOutlineLevels = 9;
            pdfSaveOptions.OutlineOptions.CreateMissingOutlineLevels = true;
            pdfSaveOptions.OutlineOptions.ExpandedOutlineLevels = 9;
              
            pdfSaveOptions.SaveFormat = SaveFormat.Pdf;

            doc.Save(MyDir + "CreateMissingOutlineLevels_OUT.pdf", pdfSaveOptions);

            //Bind pdf with Aspose PDF
            PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
            bookmarkEditor.BindPdf(MyDir + "CreateMissingOutlineLevels_OUT.pdf");

            //Get all bookmarks from the document
            Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

            Assert.AreEqual(9, bookmarks.Count);
        }

        //Note: For validation result, you can add some shapes to the document and assert, that the DML shapes are render correctly
        [Test]
        public void DrawingMl()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.DmlRenderingMode = DmlRenderingMode.DrawingML;

            doc.Save(MyDir + "DrawingMl_OUT.pdf", pdfSaveOptions);
        }

        [Test]
        public void WithoutUpdateFields()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.UpdateFields = false;

            doc.Save(MyDir + "UpdateFields_False_OUT.pdf", pdfSaveOptions);

            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(MyDir + "UpdateFields_False_OUT.pdf");

            //Get text fragment by search string
            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber("Page  of");
            pdfDocument.Pages.Accept(textFragmentAbsorber);

            //Assert that fields are not updated
            Assert.AreEqual("Page  of", textFragmentAbsorber.TextFragments[1].Text);
        }

        [Test]
        public void WithUpdateFields()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.UpdateFields = true;

            doc.Save(MyDir + "UpdateFields_False_OUT.pdf", pdfSaveOptions);

            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(MyDir + "UpdateFields_False_OUT.pdf");

            //Get text fragment by search string
            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber("Page 1 of 2");
            pdfDocument.Pages.Accept(textFragmentAbsorber);

            //Assert that fields are updated
            Assert.AreEqual("Page 1 of 2", textFragmentAbsorber.TextFragments[1].Text);
        }
    }
}
