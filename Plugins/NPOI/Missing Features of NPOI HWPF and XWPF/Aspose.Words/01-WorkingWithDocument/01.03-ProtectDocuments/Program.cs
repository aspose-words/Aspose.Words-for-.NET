using Aspose.Words;

namespace _01._03_ProtectDocuments
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document("../../data/document.doc");
            doc.Protect(ProtectionType.ReadOnly);

            doc.Save("ProtectDocuments.docx");
        }
    }
}
