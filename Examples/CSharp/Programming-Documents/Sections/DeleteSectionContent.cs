using System;
using System.Collections.Generic;
using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Sections
{
    class DeleteSectionContent
    {
        public static void Run()
        {
            //ExStart:DeleteSectionContent
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithSections();

            Document doc = new Document(dataDir + "Document.doc");
            Section section = doc.Sections[0];
            section.ClearContent();
            //ExEnd:DeleteSectionContent
            Console.WriteLine("\nSection content at 0 index deleted successfully.");
        }
    }
}
