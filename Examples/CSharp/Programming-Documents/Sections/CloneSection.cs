using System;
using System.Collections.Generic;
using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Sections
{
    class CloneSection
    {
        public static void Run()
        {
            //ExStart:CloneSection
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithSections();

            Document doc = new Document(dataDir + "Document.doc");
            Section cloneSection = doc.Sections[0].Clone();
            //ExEnd:CloneSection
            Console.WriteLine("\n0 index section clone successfully.");
        }        
    }
}
