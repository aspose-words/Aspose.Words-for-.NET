using Aspose.Words.Replacing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Find_and_Replace
{
    class FindReplaceUsingMetaCharacters
    {
        public static void Run()
        {
            /* meta-characters
            &p - paragraph break
            &b - section break
            &m - page break
            &l - manual line break
                */
            // ExStart:FindReplaceUsingMetaCharacters
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_FindAndReplace();

            // Initialize a Document.
            Document doc = new Document();

            // Use a document builder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is Line 1");
            builder.Writeln("This is Line 2");

            var findReplaceOptions = new FindReplaceOptions(); 

            doc.Range.Replace("This is Line 1&pThis is Line 2", "This is replaced line", findReplaceOptions);

            builder.MoveToDocumentEnd();
            builder.Write("This is Line 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("This is Line 2");

            doc.Range.Replace("This is Line 1&mThis is Line 2", "Page break is replaced with new text.", findReplaceOptions);

            dataDir = dataDir + "FindReplaceUsingMetaCharacters_out.docx";
            doc.Save(dataDir);
            // ExEnd:FindReplaceUsingMetaCharacters
            Console.WriteLine("\nFind and Replace text using meta-characters has done successfully.\nFile saved at " + dataDir);
        }
    }
}
