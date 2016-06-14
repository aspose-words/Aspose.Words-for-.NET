
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

            // Create a blank document.
            Document doc = new Document();

            // DocumentBuilder provides members to easily add content to a document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a new paragraph in the document with the text "Hello World!"
            builder.Writeln("Hello World!");

            // Save the document in DOCX format. The format to save as is inferred from the extension of the file name.
            // Aspose.Words supports saving any document in many more formats.
            dataDir = dataDir + "HelloWorld_out_.docx";
            doc.Save(dataDir);

            Console.WriteLine("\nNew document created successfully.\nFile saved at " + dataDir);
        }
    }
}
