using System;
using System.Collections.Generic;
using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Sections
{
    class CopySection
    {
        public static void Run()
        {
            //ExStart:CopySection
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithSections();

            Document srcDoc = new Document(dataDir + "Document.doc");
            Document dstDoc = new Document();

            Section sourceSection = srcDoc.Sections[0];
            Section newSection = (Section)dstDoc.ImportNode(sourceSection, true);
            dstDoc.Sections.Add(newSection);
            dataDir = dataDir + "Document.Copy_out_.doc";
            dstDoc.Save(dataDir);
            //ExEnd:CopySection
            Console.WriteLine("\nSection copied successfully.\nFile saved at " + dataDir);
        }        
    }
}
