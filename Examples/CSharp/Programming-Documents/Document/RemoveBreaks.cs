//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;

namespace CSharp.Programming_Documents.Working_With_Document
{
    class RemoveBreaks
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();

            // Open the document.
            Document doc = new Document(dataDir + "TestFile.doc");

            // Remove the page and section breaks from the document.
            // In Aspose.Words section breaks are represented as separate Section nodes in the document.
            // To remove these separate sections the sections are combined.
            RemovePageBreaks(doc);
            RemoveSectionBreaks(doc);

            // Save the document.
            doc.Save(dataDir + "TestFile Out.doc");

            Console.WriteLine("\nRemoved breaks from the document successfully.\nFile saved at " + dataDir + "TestFile Out.doc");
        }

        private static void RemovePageBreaks(Document doc)
        {
            // Retrieve all paragraphs in the document.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            // Iterate through all paragraphs
            foreach (Paragraph para in paragraphs)
            {
                // If the paragraph has a page break before set then clear it.
                if (para.ParagraphFormat.PageBreakBefore)
                    para.ParagraphFormat.PageBreakBefore = false;

                // Check all runs in the paragraph for page breaks and remove them.
                foreach (Run run in para.Runs)
                {
                    if (run.Text.Contains(ControlChar.PageBreak))
                        run.Text = run.Text.Replace(ControlChar.PageBreak, string.Empty);
                }

            }

        }
        
        private static void RemoveSectionBreaks(Document doc)
        {
            // Loop through all sections starting from the section that precedes the last one 
            // and moving to the first section.
            for (int i = doc.Sections.Count - 2; i >= 0; i--)
            {
                // Copy the content of the current section to the beginning of the last section.
                doc.LastSection.PrependContent(doc.Sections[i]);
                // Remove the copied section.
                doc.Sections[i].Remove();
            }
        }
    }
}
