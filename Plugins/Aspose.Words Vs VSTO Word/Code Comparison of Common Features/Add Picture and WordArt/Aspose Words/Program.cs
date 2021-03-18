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

            string filePath = @"..\..\..\..\..\Sample Files\";

            // Insert a shape into the document with an image taken from a file in the local file system.
            builder.InsertImage(filePath + "Logo.jpg");

            // Insert a shape which contains WordArt with customized text.
            Shape wordArtShape = new Shape(doc, ShapeType.TextCurve);
            wordArtShape.Width = 480;
            wordArtShape.Height = 24;
            wordArtShape.FillColor = Color.Orange;
            wordArtShape.StrokeColor = Color.Red;
            wordArtShape.WrapType = WrapType.Inline;

            wordArtShape.TextPath.FontFamily = "Arial";
            wordArtShape.TextPath.Bold = true;
            wordArtShape.TextPath.Italic = true;
            wordArtShape.TextPath.Text = "Hello World! This text is bold, and italic.";

            doc.FirstSection.Body.LastParagraph.AppendChild(wordArtShape);

            doc.Save("Add Picture and WordArt.docx");
        }
    }
}
