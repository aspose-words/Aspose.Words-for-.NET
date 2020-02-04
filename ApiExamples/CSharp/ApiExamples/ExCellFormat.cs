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
            //ExEnd
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
            //ExEnd
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

            using (MemoryStream dstStream = new MemoryStream()) builder.Document.Save(dstStream, SaveFormat.Docx);

            Table table = (Table) builder.Document.GetChild(NodeType.Table, 0, true);
            Cell cell = table.Rows[0].Cells[0];

            Assert.AreEqual(5, cell.CellFormat.LeftPadding);
            Assert.AreEqual(10, cell.CellFormat.TopPadding);
            Assert.AreEqual(40, cell.CellFormat.RightPadding);
            Assert.AreEqual(50, cell.CellFormat.BottomPadding);
        }
    }
}