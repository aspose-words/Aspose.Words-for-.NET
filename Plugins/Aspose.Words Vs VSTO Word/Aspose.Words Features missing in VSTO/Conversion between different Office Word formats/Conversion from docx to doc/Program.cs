using Aspose.Words;

namespace Conversion_from_docx_to_doc
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load a DOCX document from the local file system.
            string MyDir = @"..\..\..\..\..\Sample Files\";
            Document doc = new Document(MyDir + "MyDocument.docx");

            // Save the document to the DOC format in a different file in the local file system.
            doc.Save(MyDir + "Converted.doc", SaveFormat.Doc);
        }
    }
}
