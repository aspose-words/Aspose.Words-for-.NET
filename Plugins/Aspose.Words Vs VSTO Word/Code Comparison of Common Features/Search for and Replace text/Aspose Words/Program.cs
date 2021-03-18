using Aspose.Words;
using Aspose.Words.Replacing;

namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");

            FindReplaceOptions options = new FindReplaceOptions
            {
                MatchCase = false,
                FindWholeWordsOnly = true
            };
            
            doc.Range.Replace("Hello world!", "Greetings planet!", options);

            doc.Save("Search for and Replace text.docx");
        }
    }
}
