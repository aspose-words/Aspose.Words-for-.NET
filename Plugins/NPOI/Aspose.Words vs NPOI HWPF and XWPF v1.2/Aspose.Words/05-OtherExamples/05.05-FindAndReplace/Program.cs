using Aspose.Words;
using System;
using Aspose.Words.Replacing;

namespace _05._05_FindAndReplace
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Hello World");

            FindReplaceOptions options = new FindReplaceOptions
            {
                MatchCase = false,
                FindWholeWordsOnly = true
            };

            doc.Range.Replace("Hello", "Hallow", options);

            String text = doc.Range.Text;

            System.Console.WriteLine(text);
            System.Console.ReadKey();
        }
    }
}
