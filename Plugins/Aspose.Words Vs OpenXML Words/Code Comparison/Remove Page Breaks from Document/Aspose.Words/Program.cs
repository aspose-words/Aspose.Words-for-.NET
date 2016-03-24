using System;
using System.Collections.Generic;
using System.Text;

namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document("Test.docx");
            RemovePageBreaks(doc);
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
    }
}

