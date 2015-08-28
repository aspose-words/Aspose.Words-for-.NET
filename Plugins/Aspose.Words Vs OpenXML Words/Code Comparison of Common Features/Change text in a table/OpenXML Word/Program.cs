// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXML_Word
{
    class Program
    {
        static void Main(string[] args)
        {
            ChangeTextInCell("Change text in a table.doc", "The text from the OpenXML API example");
        }
        // Change the text in a table in a word processing document.
        public static void ChangeTextInCell(string filepath, string txt)
        {
            // Use the file name and path passed in as an argument to 
            // open an existing document.            
            using (WordprocessingDocument doc =
                WordprocessingDocument.Open(filepath, true))
            {
                // Find the first table in the document.
                Table table =
                    doc.MainDocumentPart.Document.Body.Elements<Table>().First();

                // Find the second row in the table.
                TableRow row = table.Elements<TableRow>().ElementAt(1);

                // Find the third cell in the row.
                TableCell cell = row.Elements<TableCell>().ElementAt(2);

                // Find the first paragraph in the table cell.
                Paragraph p = cell.Elements<Paragraph>().First();

                // Find the first run in the paragraph.
                Run r = p.Elements<Run>().First();

                // Set the text for the run.
                Text t = r.Elements<Text>().First();
                t.Text = txt;
            }
        }
    }
}
