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
            AlwaysCompressMetafiles(dataDir);
            SavePictureBullet(dataDir);
        }

        public static void EncryptDocumentWithPassword(string dataDir)
        {
            //ExStart:EncryptDocumentWithPassword
            Document doc = new Document(dataDir + "Document.docx");
            DocSaveOptions docSaveOptions = new DocSaveOptions();
            docSaveOptions.Password = "password";
            dataDir = dataDir + "Document.Password_out.docx";
            doc.Save(dataDir, docSaveOptions);
            //ExEnd:EncryptDocumentWithPassword
            Console.WriteLine("\nThe password of document is set using RC4 encryption method. \nFile saved at " + dataDir);
        }

        public static void AlwaysCompressMetafiles(string dataDir)
        {
            //ExStart:AlwaysCompressMetafiles
            Document doc = new Document(dataDir + "Document.doc");
            DocSaveOptions saveOptions = new DocSaveOptions();

            saveOptions.AlwaysCompressMetafiles = false;
            doc.Save(dataDir + "SmallMetafilesUncompressed.doc", saveOptions);
            //ExEnd:AlwaysCompressMetafiles
            Console.WriteLine("\nThe document is saved with AlwaysCompressMetafiles setting to false. \nFile saved at " + dataDir);
        }

        public static void SavePictureBullet(string dataDir)
        {
            //ExStart:SavePictureBullet
            Document doc = new Document(dataDir + "in.doc");
            DocSaveOptions saveOptions = (DocSaveOptions)SaveOptions.CreateSaveOptions(SaveFormat.Doc);
            saveOptions.SavePictureBullet = false;
            doc.Save(dataDir + "out.doc", saveOptions);
            //ExEnd:SavePictureBullet
            Console.WriteLine("\nThe document is saved with SavePictureBullet setting to false. \nFile saved at " + dataDir);
        }
    }
}
