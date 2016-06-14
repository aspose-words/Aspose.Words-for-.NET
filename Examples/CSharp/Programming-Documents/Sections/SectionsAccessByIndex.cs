using System;
using System.Collections.Generic;
using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Sections
{
    class SectionsAccessByIndex
    {
        public static void Run()
        {
            //ExStart:SectionsAccessByIndex
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithSections();
            Document doc = new Document(dataDir + "Document.doc");
            Section section = doc.Sections[0];
            section.PageSetup.LeftMargin = 90; // 3.17 cm
            section.PageSetup.RightMargin = 90; // 3.17 cm
            section.PageSetup.TopMargin = 72; // 2.54 cm
            section.PageSetup.BottomMargin = 72; // 2.54 cm
            section.PageSetup.HeaderDistance = 35.4; // 1.25 cm
            section.PageSetup.FooterDistance = 35.4; // 1.25 cm
            section.PageSetup.TextColumns.Spacing = 35.4; // 1.25 cm
            //ExEnd:SectionsAccessByIndex
            Console.WriteLine("\nSection at 0 index have text " + "'" + section.GetText() + "'");
        }
    }
}
