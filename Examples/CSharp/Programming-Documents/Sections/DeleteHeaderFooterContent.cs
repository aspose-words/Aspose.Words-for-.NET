using System;
using System.Collections.Generic;
using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Sections
{
    class DeleteHeaderFooterContent
    {
        public static void Run()
        {
            //ExStart:DeleteHeaderFooterContent
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithSections();

            Document doc = new Document(dataDir + "Document.doc");
            Section section = doc.Sections[0];
            section.ClearHeadersFooters();
            //ExEnd:DeleteHeaderFooterContent
            Console.WriteLine("\nHeader and footer content of 0 index deleted successfully.");
        }
    }
}
