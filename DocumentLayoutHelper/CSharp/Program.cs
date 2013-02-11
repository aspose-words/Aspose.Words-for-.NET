//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;

using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Tables;

namespace DocumentLayoutHelper
{
    class Program
    {
        static void Main(string[] args)
        {      
            string dataDir = Path.GetFullPath("../../Data/");

            Document doc = new Document(dataDir + "TestFile.docx");
           
            // This sample introduces the RenderedDocument class and other related classes which provide an API wrapper for 
            // the LayoutEnumerator. This allows you to access the layout entities of a document using a DOM style API.
            
            // Create a new RenderedDocument class from a Document object.
            RenderedDocument layoutDoc = new RenderedDocument(doc);

            // The following examples demonstrate how to use the wrapper API. 
            // This snippet returns the third line of the first page and prints the line of text to the console.
            RenderedLine line = layoutDoc.Pages[0].Columns[0].Lines[2];
            Console.WriteLine("Line: " + line.Text);

            // With a rendered line the original paragraph in the document object model can be returned.
            Paragraph para = line.Paragraph;
            Console.WriteLine("Paragraph text: " + para.Range.Text);

            // Retrieve all the text that appears of the first page in plain text format (including headers and footers).
            string pageText = layoutDoc.Pages[0].Text;
            Console.WriteLine();

            // Loop through each page in the document and print how many lines appear on each page.
            foreach (RenderedPage page in layoutDoc.Pages)
            {
                LayoutCollection<LayoutEntity> lines = page.GetChildEntities(LayoutEntityType.Line, true);
                Console.WriteLine("Page {0} has {1} lines.", page.PageIndex, lines.Count);
            }

            // This method provides a reverse lookup of lines for a given paragraph.
            Console.WriteLine();
            Console.WriteLine("The lines of the second paragraph:");
            foreach (RenderedLine paragraphLine in layoutDoc.GetLinesOfParagraph(doc.FirstSection.Body.Paragraphs[1]))
            {
                Console.WriteLine(string.Format("\"{0}\"", paragraphLine.Text.Trim()));
                Console.WriteLine(paragraphLine.Rectangle.ToString());
                Console.WriteLine();
            }
        }
    }
}
