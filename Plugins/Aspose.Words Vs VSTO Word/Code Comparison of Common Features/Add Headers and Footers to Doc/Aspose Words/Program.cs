using Aspose.Words;
using Aspose.Words.Fields;

namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a blank document and a document builder, which we will use to populate the document with content.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The document builder's cursor is currently in the body of the first section.
            // Move the cursor to the primary header of that section.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Add text to the primary header.
            builder.Write("Hello world! This is the primary header.");

            // Move the document builder to the primary footer.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

            // Insert a DATE field into the primary footer to display the current date.
            builder.InsertField(FieldType.FieldDate, true);

            // Move the cursor back into the body, to the end of the first paragraph.
            builder.MoveTo(doc.FirstSection.Body.FirstParagraph);
            builder.Write("Hello world! This is the body of the first section.");

            doc.Save("Add Headers and Footers to Doc.docx");
        }
    }
}
