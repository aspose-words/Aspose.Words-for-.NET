// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExCellFormat : ApiExampleBase
    {
        [Test]
        public void VerticalMerge()
        {
            //ExStart
            //ExFor:DocumentBuilder.EndRow
            //ExFor:CellMerge
            //ExFor:CellFormat.VerticalMerge
            //ExSummary:Shows how to merge table cells vertically.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a cell into the first column of the first row.
            // This cell will be the first in a range of vertically merged cells.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.First;
            builder.Write("Text in merged cells.");

            // Insert a cell into the second column of the first row, then end the row.
            // Also, configure the builder to disable vertical merging in created cells.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Text in unmerged cell.");
            builder.EndRow();

            // Insert a cell into the first column of the second row. 
            // Instead of adding text contents, we will merge this cell with the first cell that we added directly above.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.Previous;

            // Insert another independent cell in the second column of the second row.
            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Text in unmerged cell.");
            builder.EndRow();
            builder.EndTable();

            doc.Save(ArtifactsDir + "CellFormat.VerticalMerge.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "CellFormat.VerticalMerge.docx");
            Table table = doc.FirstSection.Body.Tables[0];

            Assert.That(table.Rows[0].Cells[0].CellFormat.VerticalMerge, Is.EqualTo(CellMerge.First));
            Assert.That(table.Rows[1].Cells[0].CellFormat.VerticalMerge, Is.EqualTo(CellMerge.Previous));
            Assert.That(table.Rows[0].Cells[0].GetText().Trim('\a'), Is.EqualTo("Text in merged cells."));
            Assert.That(table.Rows[1].Cells[0].GetText(), Is.Not.EqualTo(table.Rows[0].Cells[0].GetText()));
        }

        [Test]
        public void HorizontalMerge()
        {
            //ExStart
            //ExFor:CellMerge
            //ExFor:CellFormat.HorizontalMerge
            //ExSummary:Shows how to merge table cells horizontally.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a cell into the first column of the first row.
            // This cell will be the first in a range of horizontally merged cells.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Text in merged cells.");

            // Insert a cell into the second column of the first row. Instead of adding text contents,
            // we will merge this cell with the first cell that we added directly to the left.
            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.EndRow();

            // Insert two more unmerged cells to the second row.
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.InsertCell();
            builder.Write("Text in unmerged cell.");
            builder.InsertCell();
            builder.Write("Text in unmerged cell.");
            builder.EndRow();
            builder.EndTable();

            doc.Save(ArtifactsDir + "CellFormat.HorizontalMerge.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "CellFormat.HorizontalMerge.docx");
            Table table = doc.FirstSection.Body.Tables[0];

            Assert.That(table.Rows[0].Cells.Count, Is.EqualTo(1));
            Assert.That(table.Rows[0].Cells[0].CellFormat.HorizontalMerge, Is.EqualTo(CellMerge.None));
            Assert.That(table.Rows[0].Cells[0].GetText().Trim('\a'), Is.EqualTo("Text in merged cells."));
        }

        [Test]
        public void Padding()
        {
            //ExStart
            //ExFor:CellFormat.SetPaddings
            //ExSummary:Shows how to pad the contents of a cell with whitespace.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set a padding distance (in points) between the border and the text contents
            // of each table cell we create with the document builder. 
            builder.CellFormat.SetPaddings(5, 10, 40, 50);

            // Create a table with one cell whose contents will have whitespace padding.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. " +
                          "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.");

            doc.Save(ArtifactsDir + "CellFormat.Padding.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "CellFormat.Padding.docx");

            Table table = doc.FirstSection.Body.Tables[0];
            Cell cell = table.Rows[0].Cells[0];

            Assert.That(cell.CellFormat.LeftPadding, Is.EqualTo(5));
            Assert.That(cell.CellFormat.TopPadding, Is.EqualTo(10));
            Assert.That(cell.CellFormat.RightPadding, Is.EqualTo(40));
            Assert.That(cell.CellFormat.BottomPadding, Is.EqualTo(50));
        }
    }
}