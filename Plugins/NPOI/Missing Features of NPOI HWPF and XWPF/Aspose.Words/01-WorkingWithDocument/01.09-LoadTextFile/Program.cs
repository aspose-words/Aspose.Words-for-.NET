using Aspose.Words;

namespace _01._09_LoadTextFile
{
    class Program
    {
        static void Main(string[] args)
        {
            // The encoding of the text file is automatically detected.
            Document doc = new Document("../../data/LoadTxt.txt");

            // Save as any Aspose.Words supported format, such as DOCX.
            doc.Save("LoadTextFile.docx");
        }
    }
}
