using Aspose.Words;

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
            
            doc.Save("Add Picture.docx");
        }
    }
}
