//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using System.Reflection;

using Aspose.Words;

namespace RemoveBreaks
{
    class Program
    {
        public static void Main(string[] args)
        {
            // The sample infrastructure.
            string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar;
            string dataDir = new Uri(new Uri(exeDir), @"../../Data/").LocalPath;

            // Open the document.
            Document doc = new Document(dataDir + "TestFile.doc");
            
            // Remove the page and section breaks from the document.
            // In Aspose.Words section breaks are represented as separate Section nodes in the document.
            // To remove these separate sections the sections are combined.
            RemovePageBreaks(doc);
            RemoveSectionBreaks(doc);

            // Save the document.
            doc.Save(dataDir + "TestFile Out.doc");
        }

        //ExStart
        //ExFor:ControlChar.PageBreak
        //ExId:RemoveBreaks_Pages
        //ExSummary:Removes all page breaks from the document.
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
        //ExEnd


        //ExStart
        //ExId:RemoveBreaks_Sections
        //ExSummary:Combines all sections in the document into one.
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
        //ExEnd
    }
}
