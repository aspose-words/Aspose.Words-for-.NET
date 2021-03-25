using Aspose.Words;
namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");

            // Insert a primary header and a primary footer.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("This is the primary header.");

            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write("This is the primary footer.");

            // Remove all headers and footers from the document.
            foreach (Section section in doc)
                foreach (HeaderFooter headerFooter in section.HeadersFooters)
                    headerFooter.Remove();

            doc.Save("Removing Header and Footer.docx");
        }
    }
}
