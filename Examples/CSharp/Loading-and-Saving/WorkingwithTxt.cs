using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using System.Text;
using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class WorkingWithTxt
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            SaveAsTxt(dataDir);
            AddBidiMarks(dataDir);
            DetectNumberingWithWhitespaces(dataDir);
            HandleSpacesOptions(dataDir);
        }

        public static void SaveAsTxt(string dataDir)
        {
            //ExStart:SaveAsTxt
            Document doc = new Document(dataDir + "Document.doc");
            dataDir = dataDir + "Document.ConvertToTxt_out.txt";
            doc.Save(dataDir);
            //ExEnd:SaveAsTxt
            Console.WriteLine("\nDocument saved as TXT.\nFile saved at " + dataDir);
        }

        public static void AddBidiMarks(string dataDir)
        {
            //ExStart:AddBidiMarks
            Document doc = new Document(dataDir + "Input.docx");
            TxtSaveOptions saveOptions = new TxtSaveOptions();
            saveOptions.AddBidiMarks = false;

            dataDir = dataDir + "Document.AddBidiMarks_out.txt";
            doc.Save(dataDir, saveOptions);
            //ExEnd:AddBidiMarks
            Console.WriteLine("\nAdd bi-directional marks set successfully.\nFile saved at " + dataDir);
        }
         
        public static void DetectNumberingWithWhitespaces(string dataDir)
        {
            //ExStart:DetectNumberingWithWhitespaces
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            loadOptions.DetectNumberingWithWhitespaces = false;

            Document doc = new Document(dataDir + "LoadTxt.txt", loadOptions);

            dataDir = dataDir + "DetectNumberingWithWhitespaces_out.docx";
            doc.Save(dataDir);
            //ExEnd:DetectNumberingWithWhitespaces
            Console.WriteLine("\nDetect number with whitespaces successfully.\nFile saved at " + dataDir);
        }

        public static void HandleSpacesOptions(string dataDir)
        {
            //ExStart:HandleSpacesOptions
            TxtLoadOptions loadOptions = new TxtLoadOptions();
             
            loadOptions.LeadingSpacesOptions = TxtLeadingSpacesOptions.Trim;
            loadOptions.TrailingSpacesOptions = TxtTrailingSpacesOptions.Trim;
            Document doc = new Document(dataDir + "LoadTxt.txt", loadOptions);

            dataDir = dataDir + "HandleSpacesOptions_out.docx";
            doc.Save(dataDir);
            //ExEnd:HandleSpacesOptions
            Console.WriteLine("\nTrim leading and trailing spaces while importing text document.\nFile saved at " + dataDir);
        }
    }
}
