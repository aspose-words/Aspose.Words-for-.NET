//ExStart:AsposeWordsDockerTest
using System;
using Aspose.Words;

namespace Docker
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            // Create document and save it in all available formats.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello from Aspose.Words!!!");

            foreach (SaveFormat sf in Enum.GetValues(typeof(SaveFormat)))
            {
                if (sf != SaveFormat.Unknown)
                {
                    try
                    {
                        doc.Save($"out{FileFormatUtil.SaveFormatToExtension(sf)}", sf);
                        Console.WriteLine("Saving {0}\t\t[OK]", sf);
                    }
                    catch
                    {
                        Console.WriteLine("Saving {0}\t\t[FAILED]", sf);
                    }
                }
            }
        }
    }
}
//ExEnd:AsposeWordsDockerTest