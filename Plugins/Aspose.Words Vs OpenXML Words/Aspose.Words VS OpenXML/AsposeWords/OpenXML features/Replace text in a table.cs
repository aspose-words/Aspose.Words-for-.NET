// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class ReplaceTextInATable : TestUtil
    {
        [Test]
        public void ReplaceText()
        {
            File.Copy(MyDir + "Replace text.docx", ArtifactsDir + "Replace text - OpenXML.docx", true);

            // Use the file name and path passed in as an argument to open an existing document.
            using WordprocessingDocument doc = WordprocessingDocument.Open(ArtifactsDir + "Replace text - OpenXML.docx", true);

            // Get the main document part.
            MainDocumentPart mainPart = doc.MainDocumentPart;
            if (mainPart?.Document?.Body == null)
                throw new InvalidOperationException("The document does not contain a valid body.");

            // Find the first table in the document.
            Table table = mainPart.Document.Body.Elements<Table>().FirstOrDefault();
            // Find the second row in the table.
            TableRow row = table.Elements<TableRow>().ElementAtOrDefault(1);
            // Find the third cell in the row.
            TableCell cell = row.Elements<TableCell>().ElementAtOrDefault(2);
            // Find the first paragraph in the table cell.
            Paragraph paragraph = cell.Elements<Paragraph>().FirstOrDefault();
            // Find the first run in the paragraph.
            Run run = paragraph.Elements<Run>().FirstOrDefault();

            // Find the first text element in the run.
            Text text = run.Elements<Text>().FirstOrDefault();
            if (text == null)
            {
                // If no text element exists, create one.
                text = new Text();
                run.Append(text);
            }

            // Set the text for the run.
            text.Text = "The text from the OpenXML API example";
        }
    }
}
