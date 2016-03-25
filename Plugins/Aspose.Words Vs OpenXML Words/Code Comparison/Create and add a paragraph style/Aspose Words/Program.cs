using Aspose.Words;
namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {

            // Open the document.
            Document doc = new Document("Create and add a paragraph style.doc");

            DocumentBuilder builder = new DocumentBuilder(doc);
            // Set font formatting properties
            Aspose.Words.Font font = builder.Font;
            font.Bold = true;
            font.Color = System.Drawing.Color.Red;
            font.Italic = true;
            font.Name = "Arial";
            font.Size = 24;
            font.Spacing = 5;
            font.Underline = Underline.Double;

            // Output formatted text
            builder.MoveToDocumentEnd();
            builder.Writeln("I'm a very nice formatted string. zeeshan");

            string txt = builder.CurrentParagraph.GetText();

            doc.Save("Create and add a paragraph style.doc");
        }

    }
}
