// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class ChangeTextInATable : TestUtil
    {
        [Test]
        public void ChangeTextInATableFeature()
        {
            // Use the file name and path passed in as an argument to 
            // open an existing document.            
            using (WordprocessingDocument doc =
                WordprocessingDocument.Open(MyDir + "Change text in a table.docx", true))
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
                t.Text = "The text from the OpenXML API example";
            }
        }
    }
}
