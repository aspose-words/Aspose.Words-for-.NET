using Aspose.Words;

namespace _01._02_CloneDocuments
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document("../../data/document.doc");
            Document clone = doc.Clone();

            clone.Save("CloneDocuments.docx");
        }
    }
}
