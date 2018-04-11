using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Loading_and_Saving
{
    class Load_Options
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_QuickStart();

            LoadOptionsUpdateDirtyFields(dataDir);
            LoadAndSaveEncryptedODT(dataDir);
            VerifyODTdocument(dataDir);
        }

        public static void LoadOptionsUpdateDirtyFields(string dataDir)
        {
            // ExStart:LoadOptionsUpdateDirtyFields  
            LoadOptions lo = new LoadOptions();

            //Update the fields with the dirty attribute
            lo.UpdateDirtyFields = true;

            //Load the Word document
            Document doc = new Document(dataDir + @"input.docx", lo);
             
            //Save the document into DOCX
            doc.Save(dataDir + "output.docx", SaveFormat.Docx);
            // ExEnd:LoadOptionsUpdateDirtyFields 
            Console.WriteLine("\nUpdate the fields with the dirty attribute successfully.\nFile saved at " + dataDir);
        }

        public static void LoadAndSaveEncryptedODT(string dataDir)
        {
            // ExStart:LoadAndSaveEncryptedODT  
            Document doc = new Document(dataDir + @"encrypted.odt", new Aspose.Words.LoadOptions("password"));

            doc.Save(dataDir + "out.odt", new OdtSaveOptions("newpassword"));
            // ExEnd:LoadAndSaveEncryptedODT 
            Console.WriteLine("\nLoad and save encrypted document successfully.\nFile saved at " + dataDir);
        }

        public static void VerifyODTdocument(string dataDir)
        {
            // ExStart:VerifyODTdocument  
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(dataDir + @"encrypted.odt");
            Console.WriteLine(info.IsEncrypted);
            // ExEnd:VerifyODTdocument 
        }
    }
}
