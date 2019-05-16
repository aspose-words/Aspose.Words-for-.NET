using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Joining_Appending
{
    class AppendWithImportFormatOptions
    {
        public static void Run()
        {
            // ExStart:AppendWithImportFormatOptions
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_JoiningAndAppending();
            Document srcDoc = new Document(dataDir + "source.docx");
            Document dstDoc = new Document(dataDir + "destination.docx");

            ImportFormatOptions options = new ImportFormatOptions();
            // Specify that if numbering clashes in source and destination documents, then a numbering from the source document will be used.
            options.KeepSourceNumbering = true;
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles, options);
            // ExEnd:AppendWithImportFormatOptions
            Console.WriteLine("\nDocument appended successfully with keep source numbering option.\nFile saved at " + dataDir);
        }
    }
}
