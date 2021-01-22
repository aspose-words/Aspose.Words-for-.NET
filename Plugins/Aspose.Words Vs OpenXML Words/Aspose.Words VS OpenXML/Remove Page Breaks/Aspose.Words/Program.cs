// Copyright (c) Aspose 2002-2021. All Rights Reserved.

/*
    This project uses NuGet's Automatic Package Restore feature to 
    resolve the Aspose.Words for .NET API reference when the project is built. 
    Please visit https://docs.nuget.org/consume/nuget-faq for more information. 

    If you do not wish to use NuGet, you can manually download Aspose.Words for .NET API 
    from http://www.aspose.com/downloads, install it, and then add a reference to it to this project. 

    For any issues, questions or suggestions, please visit the Aspose Forums: https://forum.aspose.com/
*/

using Aspose.Words;

namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string fileName = FilePath + "Remove Page Breaks.docx";

            Document doc = new Document(fileName);
            RemovePageBreaks(doc);
        }

        private static void RemovePageBreaks(Document doc)
        {
            // Retrieve all paragraphs in the document.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

            foreach (Paragraph para in paragraphs)
            {
                // If the paragraph has a page break set before, then clear it.
                if (para.ParagraphFormat.PageBreakBefore)
                    para.ParagraphFormat.PageBreakBefore = false;

                // Check all runs in the paragraph for page breaks and remove them.
                foreach (Run run in para.Runs)
                    if (run.Text.Contains(ControlChar.PageBreak))
                        run.Text = run.Text.Replace(ControlChar.PageBreak, string.Empty);
            }
        }
    }
}

