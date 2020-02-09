// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
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
            //ExSummary:Creates a table with two columns with cells merged vertically in the first column.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.First;
            builder.Write("Text in merged cells.");

            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Text in one cell");
            builder.EndRow();

            builder.InsertCell();
            // This cell is vertically merged to the cell above and should be empty.
            builder.CellFormat.VerticalMerge = CellMerge.Previous;

            builder.InsertCell();
            builder.CellFormat.VerticalMerge = CellMerge.None;
            builder.Write("Text in another cell");
            builder.EndRow();
            builder.EndTable();

            doc.Save(ArtifactsDir + "CellFormat.VerticalMerge.docx");
            //ExEnd

            Document outDoc = new Document(ArtifactsDir + "CellFormat.VerticalMerge.docx");
            Table table = (Table)outDoc.GetChild(NodeType.Table, 0, true);
            Assert.AreEqual(CellMerge.First, table.Rows[0].Cells[0].CellFormat.VerticalMerge);
            Assert.AreEqual(CellMerge.Previous, table.Rows[1].Cells[0].CellFormat.VerticalMerge);

            // After the merge both cells still exist, and the one with the VerticalMerge set to "First" overlaps both of them 
            // and only that cell contains the shared text
            Assert.AreEqual("Text in merged cells.", table.Rows[0].Cells[0].GetText().Trim('\a'));
            Assert.AreNotEqual(table.Rows[0].Cells[0].GetText(), table.Rows[1].Cells[0].GetText());
        }

        [Test]
        public void HorizontalMerge()
        {
            //ExStart
            //ExFor:CellMerge
            //ExFor:CellFormat.HorizontalMerge
            //ExSummary:Creates a table with two rows with cells in the first row horizontally merged.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Text in merged cells.");

            builder.InsertCell();
            // This cell is merged to the previous and should be empty.
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;
            builder.EndRow();

            builder.InsertCell();
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.Write("Text in one cell.");

            builder.InsertCell();
            builder.Write("Text in another cell.");
            builder.EndRow();
            builder.EndTable();

            doc.Save(ArtifactsDir + "CellFormat.HorizontalMerge.docx");
            //ExEnd

            Document outDoc = new Document(ArtifactsDir + "CellFormat.HorizontalMerge.docx");
            Table table = (Table)outDoc.GetChild(NodeType.Table, 0, true);

            // Compared to the vertical merge, where both cells are still present, 
            // the horizontal merge actually removes cells with a HorizontalMerge set to "Previous" if overlapped by ones with "First"
            // Thus the first row that we inserted two cells into now has one, which is a normal cell with a HorizontalMerge of "None"
            Assert.AreEqual(1, table.Rows[0].Cells.Count);
            Assert.AreEqual(CellMerge.None, table.Rows[0].Cells[0].CellFormat.HorizontalMerge);

            Assert.AreEqual("Text in merged cells.", table.Rows[0].Cells[0].GetText().Trim('\a'));
        }

        [Test]
        public void SetCellPaddings()
        {
            //ExStart
            //ExFor:CellFormat.SetPaddings
            //ExSummary:Shows how to set paddings to a table cell.
            DocumentBuilder builder = new DocumentBuilder();

            builder.StartTable();
            builder.CellFormat.Width = 300;
            builder.CellFormat.SetPaddings(5, 10, 40, 50);

            builder.RowFormat.HeightRule = HeightRule.Exactly;
            builder.RowFormat.Height = 50;

            builder.InsertCell();
            builder.Write("Row 1, Col 1");
            //ExEnd

            using (MemoryStream dstStream = new MemoryStream())
            {
                builder.Document.Save(dstStream, SaveFormat.Docx);
                Document outDoc = new Document(dstStream);

                Table table = (Table)outDoc.GetChild(NodeType.Table, 0, true);
                Cell cell = table.Rows[0].Cells[0];

                Assert.AreEqual(5, cell.CellFormat.LeftPadding);
                Assert.AreEqual(10, cell.CellFormat.TopPadding);
                Assert.AreEqual(40, cell.CellFormat.RightPadding);
                Assert.AreEqual(50, cell.CellFormat.BottomPadding);
            }
        }
    }
}