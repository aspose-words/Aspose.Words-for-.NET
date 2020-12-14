using System;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class CloningDocument
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();

            DeepDocumentCopy(dataDir);
            DuplicateSection();
        }

        public static void DeepDocumentCopy(string dataDir)
        {
            // ExStart:CloningDocument
            Document doc = new Document(dataDir + "TestFile.doc");

            Document clone = doc.Clone();

            dataDir = dataDir + "TestFile_clone_out.doc";

            // Save the document to disk.
            clone.Save(dataDir);
            // ExEnd:CloningDocument
            Console.WriteLine("\nDocument cloned successfully.\nFile saved at " + dataDir);
        }

        public static void DuplicateSection()
        {
            //ExStart:DuplicateSection
            // Create a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is the original document before applying the clone method");

            // Clone the document.
            Document clone = doc.Clone();

            // Edit the cloned document.
            builder = new DocumentBuilder(clone);
            builder.Write("Section 1");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Write("Section 2");

            // This shows what is in the document originally. The document has two sections.
            Assert.AreEqual("Section 1\x000cSection 2", clone.GetText().Trim());

            // Duplicate the last section and append the copy to the end of the document.
            int lastSectionIdx = clone.Sections.Count - 1;
            Section newSection = clone.Sections[lastSectionIdx].Clone();
            clone.Sections.Add(newSection);

            // Check what the document contains after we changed it.
            Assert.AreEqual("Section 1\x000cSection 2", clone.GetText().Trim());
            //ExEnd:DuplicateSection
        }
    }
}
