using System;
using System.Collections.Generic;
using Aspose.Words;
using System.IO;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class OpenEncryptedDocument
    {
        public static void Run()
        {
            //ExStart:OpenEncryptedDocument      
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();
            
            //Loads encrypted document.
            Document doc = new Document(dataDir + "LoadEncrypted.docx", new LoadOptions("aspose"));

            //ExEnd:OpenEncryptedDocument

            Console.WriteLine("\nEncrypted document loaded successfully.");

        }
    }
}
