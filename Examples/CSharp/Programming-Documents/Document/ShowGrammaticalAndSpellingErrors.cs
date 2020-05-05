using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class ShowGrammaticalAndSpellingErrors
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();

            // ExStart: ShowGrammaticalAndSpellingErrors
            Document doc = new Document(dataDir + "Document.doc");

            doc.ShowGrammaticalErrors = true;
            doc.ShowSpellingErrors = true;

            doc.Save(dataDir + "Document.ShowErrorsInDocument_out.docx");
            // ExEnd: ShowGrammaticalAndSpellingErrors
            Console.WriteLine("\nDocument saved successfully.\nFile saved at " + dataDir);
        }
    }
}
