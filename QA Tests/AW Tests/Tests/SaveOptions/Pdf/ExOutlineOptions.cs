using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace QA_Tests.Tests.SaveOptions.Pdf
{
    [TestFixture]
    internal class ExOutlineOptions : QaTestsBase
    {
        [Test]
        public void CreateMissingOutlineLevels()
        {
            //ExStart
            //ExFor:Saving.PdfSaveOptions.OutlineOptions.CreateMissingOutlineLevels
            //ExSummary:Shows how to create missing outline levels saving the document in pdf
            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);

            // Creating TOC entries
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            builder.Writeln("Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading4;

            builder.Writeln("Heading 1.1.1.1");
            builder.Writeln("Heading 1.1.1.2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading9;

            builder.Writeln("Heading 1.1.1.1.1.1.1.1.1");
            builder.Writeln("Heading 1.1.1.1.1.1.1.1.2");

            //Create "PdfSaveOptions" with some mandatory parameters
            //"HeadingsOutlineLevels" specifies how many levels of headings to include in the document outline
            //"CreateMissingOutlineLevels" determining whether or not to create missing heading levels
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            pdfSaveOptions.OutlineOptions.HeadingsOutlineLevels = 9;
            pdfSaveOptions.OutlineOptions.CreateMissingOutlineLevels = true;
            pdfSaveOptions.SaveFormat = SaveFormat.Pdf;

            doc.Save(MyDir + "CreateMissingOutlineLevels.pdf", pdfSaveOptions);
            //ExEnd
        }
    }
}
