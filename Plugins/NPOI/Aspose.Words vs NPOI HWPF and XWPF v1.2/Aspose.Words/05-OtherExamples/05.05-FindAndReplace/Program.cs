using Aspose.Words;
using System;

namespace _05._05_FindAndReplace
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Hello World");

            doc.Range.Replace("Hello", "Hallow", false, true);

            String text = doc.Range.Text;

            System.Console.WriteLine(text);
            System.Console.ReadKey();
        }
    }
}
