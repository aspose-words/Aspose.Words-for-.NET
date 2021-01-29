using Aspose.Words;

namespace _01._01_AppendDocuments
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc1 = new Document("../../data/doc1.doc");
            Document doc2 = new Document("../../data/doc2.doc");

            Document doc3 = doc1.Clone();
            doc3.AppendDocument(doc2, ImportFormatMode.KeepSourceFormatting);
            doc3.Save("AppendDocuments.docx");
        }
    }
}
