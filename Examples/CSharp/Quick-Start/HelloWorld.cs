
using System.IO;

using Aspose.Words;
using System;

namespace Aspose.Words.Examples.CSharp.Quick_Start
{
    class HelloWorld
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_QuickStart();

            //ExStart:HelloWorld
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Hello, world!");
            //ExEnd:HelloWorld

            // Save the document in DOCX format. The format to save as is inferred from the extension of the file name.
            // Aspose.Words supports saving any document in many more formats.
            dataDir = dataDir + "HelloWorld_out.docx";
            doc.Save(dataDir);

            Console.WriteLine("\nNew document created successfully.\nFile saved at " + dataDir);
        }
    }
}
