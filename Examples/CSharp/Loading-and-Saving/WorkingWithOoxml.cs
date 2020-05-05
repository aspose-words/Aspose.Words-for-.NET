using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class WorkingWithOoxml
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            EncryptDocxWithPassword(dataDir);
            SetOOXMLCompliance(dataDir);
            UpdateLastSavedTimeProperty(dataDir);
            KeepLegacyControlChars(dataDir);
            SetCompressionLevel(dataDir);
        }

        public static void EncryptDocxWithPassword(string dataDir)
        {
            //ExStart:EncryptDocxWithPassword
            Document doc = new Document(dataDir + "Document.doc");
            OoxmlSaveOptions ooxmlSaveOptions = new OoxmlSaveOptions();
            ooxmlSaveOptions.Password = "password";
            dataDir = dataDir + "Document.Password_out.docx";
            doc.Save(dataDir, ooxmlSaveOptions);
            //ExEnd:EncryptDocxWithPassword
            Console.WriteLine("\nThe password of document is set using ECMA376 Standard encryption algorithm.\nFile saved at " + dataDir);
        }

        public static void SetOOXMLCompliance(string dataDir)
        {
            //ExStart:SetOOXMLCompliance
            Document doc = new Document(dataDir + "Document.doc");
            
            // Set Word2016 version for document
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2016);

            //Set the Strict compliance level. 
            OoxmlSaveOptions ooxmlSaveOptions = new OoxmlSaveOptions
            {
                Compliance = OoxmlCompliance.Iso29500_2008_Strict,
                SaveFormat = SaveFormat.Docx
            };
            dataDir = dataDir + "Document.Iso29500_2008_Strict_out.docx";
            doc.Save(dataDir, ooxmlSaveOptions);
            //ExEnd:SetOOXMLCompliance
            Console.WriteLine("\nDocument is saved with ISO/IEC 29500:2008 Strict compliance level.\nFile saved at " + dataDir);
        }

        public static void UpdateLastSavedTimeProperty(String dataDir)
        {
            // ExStart:UpdateLastSavedTimeProperty
            Document doc = new Document(dataDir + "Document.doc");

            OoxmlSaveOptions ooxmlSaveOptions = new OoxmlSaveOptions();
            ooxmlSaveOptions.UpdateLastSavedTimeProperty = true;

            dataDir = dataDir + "UpdateLastSavedTimeProperty_out.docx";

            // Save the document to disk.
            doc.Save(dataDir, ooxmlSaveOptions);
            // ExEnd:UpdateLastSavedTimeProperty
            Console.WriteLine("\nUpdated Last Saved Time Property successfully.\nFile saved at " + dataDir);
        }

        public static void KeepLegacyControlChars(String dataDir)
        {
            // ExStart:KeepLegacyControlChars
            Document doc = new Document(dataDir + "Document.doc");

            OoxmlSaveOptions so = new OoxmlSaveOptions(SaveFormat.FlatOpc);
            so.KeepLegacyControlChars = true;

            dataDir = dataDir + "Document_out.docx";
            // Save the document to disk.
            doc.Save(dataDir, so);

            // ExEnd:KeepLegacyControlChars
            Console.WriteLine("\nUpdated Last Saved With Keeping Legacy Control Chars Successfully.\nFile saved at " + dataDir);
        }

        public static void SetCompressionLevel(string dataDir)
        {
            // ExStart:SetCompressionLevel
            Document doc = new Document(dataDir + "Document.doc");

            OoxmlSaveOptions so = new OoxmlSaveOptions(SaveFormat.Docx);
            so.CompressionLevel = CompressionLevel.SuperFast;

            // Save the document to disk.
            doc.Save(dataDir + "SetCompressionLevel_out.docx", so);

            // ExEnd:SetCompressionLevel
            Console.WriteLine("\nDocument save with a Compression Level Successfully.\nFile saved at " + dataDir);
            doc.Save("out.docx", so);
        }
    }
}
