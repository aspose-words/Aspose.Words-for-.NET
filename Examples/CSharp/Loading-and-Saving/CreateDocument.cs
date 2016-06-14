using System;
using System.Collections.Generic;
using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class CreateDocument
    {
        public static void Run()
        {
            //ExStart:CreateDocument            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            // Initialize a Document.
            Document doc = new Document();
            
            // Use a document builder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello World!");  
        
            dataDir  = dataDir + "CreateDocument_out_.docx";
            // Save the document to disk.
            doc.Save(dataDir);
           
            //ExEnd:CreateDocument

            Console.WriteLine("\nDocument created successfully.\nFile saved at " + dataDir);

        }
    }
}
