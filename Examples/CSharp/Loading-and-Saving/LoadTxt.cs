using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using System.Diagnostics;
using Aspose.Words.MailMerging;
using Aspose.Words.Saving;
using System.Text;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class LoadTxt
    {
        public static void Run()
        {
            //ExStart:LoadTxt
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();
            
            // The encoding of the text file is automatically detected.
            Document doc = new Document(dataDir + "LoadTxt.txt");

            // Save as any Aspose.Words supported format, such as DOCX.  
            dataDir = dataDir + "LoadTxt_out_.docx";
            doc.Save(dataDir);
            //ExEnd:LoadTxt
            Console.WriteLine("\nText document loaded successfully.\nFile saved at " + dataDir);
        }
    }
}
