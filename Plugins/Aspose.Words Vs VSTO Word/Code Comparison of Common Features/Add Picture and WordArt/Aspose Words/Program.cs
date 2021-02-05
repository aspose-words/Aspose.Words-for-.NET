using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a blank document and a document builder which we will use to populate the document with content.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a shape which contains and displays an image taken from a file in the local file system.
            builder.InsertImage("download.jpg");

            // Insert a shape which contains WordArt with customized text.
            Shape shape = new Shape(doc, ShapeType.TextPlainText)
            {
                WrapType = WrapType.Inline,
                Width = 480,
                Height = 24,
                FillColor = Color.Orange,
                StrokeColor = Color.Red
            };

            shape.TextPath.FontFamily = "Arial";
            shape.TextPath.Bold = true;
            shape.TextPath.Italic = true;
            shape.TextPath.Text = "Hello World! This text is bold, and italic.";

            doc.Save("Add Picture and WordArt.docx");
        }
    }
}
