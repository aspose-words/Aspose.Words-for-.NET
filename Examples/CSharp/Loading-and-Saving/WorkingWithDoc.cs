using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class WorkingWithDoc
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            EncryptDocumentWithPassword(dataDir);
        }

        public static void EncryptDocumentWithPassword(string dataDir)
        {
            //ExStart:EncryptDocumentWithPassword
            Document doc = new Document(dataDir + "Document.doc");
            DocSaveOptions docSaveOptions = new DocSaveOptions();
            docSaveOptions.Password = "password";
            dataDir = dataDir + "Document.Password_out.doc";
            doc.Save(dataDir, docSaveOptions);
            //ExEnd:EncryptDocumentWithPassword
            Console.WriteLine("\nThe password of document is set using RC4 encryption method. \nFile saved at " + dataDir);
        }
    }
}
