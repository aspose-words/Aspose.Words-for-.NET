using Aspose.Words;

namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a blank document, and a document builder which we will use to populate the document with content.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world! This is the first section.");

            // Use the document builder to start a new section on a fresh page.
            builder.InsertBreak(BreakType.SectionBreakNewPage);

            builder.Writeln("Hello again! This is the second section.");

            // If we wish to apply page setup changes to an entire document, we will need to iterate over every section.
            foreach (Section section in doc)
                section.PageSetup.PaperSize = PaperSize.EnvelopeDL;

            doc.Save("Changing Page Setup for Whole Document.docx");
        }
    }
}
