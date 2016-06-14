using System;
using System.Collections.Generic;
using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Sections
{
    class AddDeleteSection
    {
        public static void Run()
        {
            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithSections() + "Section.AddRemove.doc";
            AddSection(dataDir);
            DeleteSection(dataDir);
            DeleteAllSections(dataDir);
        }
        /// <summary>
        /// Shows how to add a section to the end of the document.
        /// </summary>
        private static void AddSection(string dataDir)
        {
            //ExStart:AddSection
            Document doc = new Document(dataDir);
            Section sectionToAdd = new Section(doc);
            doc.Sections.Add(sectionToAdd);
            //ExEnd:AddSection
            Console.WriteLine("\nSection added successfully to the end of the document.");
        }
        /// <summary>
        /// Shows how to remove a section at the specified index.
        /// </summary>
        private static void DeleteSection(string dataDir)
        {
            //ExStart:DeleteSection
            Document doc = new Document(dataDir);
            doc.Sections.RemoveAt(0);
            //ExEnd:DeleteSection
            Console.WriteLine("\nSection deleted successfully at 0 index.");
        }
        /// <summary>
        /// Shows how to remove all sections from a document.
        /// </summary>
        private static void DeleteAllSections(string dataDir)
        {
            //ExStart:DeleteAllSections
            Document doc = new Document(dataDir);
            doc.Sections.Clear();
            //ExEnd:DeleteAllSections
            Console.WriteLine("\nAll sections deleted successfully form the document.");
        }
    }
}
