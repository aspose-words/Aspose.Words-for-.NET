using System.Drawing;
using Aspose.Words;
namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document("Sample Aspose.doc");

            DocumentBuilder builder = new DocumentBuilder(doc);

            // Specify font formatting before adding text.
            Aspose.Words.Font font = builder.Font;
            font.Size = 16;
            font.Bold = true;
            font.Color = Color.Blue;
            font.Name = "Arial";
            font.Underline = Underline.Dash;

            builder.Write("Insert text");
            doc.Save("Add text Out.doc");
        }
    }
}
