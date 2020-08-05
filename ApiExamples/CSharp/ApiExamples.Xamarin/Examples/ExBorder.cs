// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExBorder : ApiExampleBase
    {
        [Test]
        public void FontBorder()
        {
            //ExStart
            //ExFor:Border
            //ExFor:Border.Color
            //ExFor:Border.LineWidth
            //ExFor:Border.LineStyle
            //ExFor:Font.Border
            //ExFor:LineStyle
            //ExFor:Font
            //ExFor:DocumentBuilder.Font
            //ExFor:DocumentBuilder.Write(String)
            //ExSummary:Shows how to insert a string surrounded by a border into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Border.Color = Color.Green;
            builder.Font.Border.LineWidth = 2.5d;
            builder.Font.Border.LineStyle = LineStyle.DashDotStroker;

            builder.Write("Text surrounded by green border.");

            doc.Save(ArtifactsDir + "Border.FontBorder.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Border.FontBorder.docx");
            Border border = doc.FirstSection.Body.FirstParagraph.Runs[0].Font.Border;

            Assert.AreEqual(Color.Green.ToArgb(), border.Color.ToArgb());
            Assert.AreEqual(2.5d, border.LineWidth);
            Assert.AreEqual(LineStyle.DashDotStroker, border.LineStyle);
        }

        [Test]
        public void ParagraphTopBorder()
        {
            //ExStart
            //ExFor:BorderCollection
            //ExFor:Border
            //ExFor:BorderType
            //ExFor:ParagraphFormat.Borders
            //ExSummary:Shows how to insert a paragraph with a top border.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Border topBorder = builder.ParagraphFormat.Borders[BorderType.Top];
            topBorder.Color = Color.Red;
            topBorder.LineWidth = 4.0d;
            topBorder.LineStyle = LineStyle.DashSmallGap;

            builder.Writeln("Text with a red top border.");

            doc.Save(ArtifactsDir + "Border.ParagraphTopBorder.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Border.ParagraphTopBorder.docx");
            Border border = doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Borders[BorderType.Top];

            Assert.AreEqual(Color.Red.ToArgb(), border.Color.ToArgb());
            Assert.AreEqual(4.0d, border.LineWidth);
            Assert.AreEqual(LineStyle.DashSmallGap, border.LineStyle);
        }

        [Test]
        public void ClearFormatting()
        {
            //ExStart
            //ExFor:Border.ClearFormatting
            //ExSummary:Shows how to remove borders from a paragraph.
            Document doc = new Document(MyDir + "Borders.docx");

            // Get the first paragraph's collection of borders
            DocumentBuilder builder = new DocumentBuilder(doc);
            BorderCollection borders = builder.ParagraphFormat.Borders;
            Assert.AreEqual(Color.Red.ToArgb(), borders[0].Color.ToArgb()); //ExSkip
            Assert.AreEqual(3.0d, borders[0].LineWidth); // ExSkip
            Assert.AreEqual(LineStyle.Single, borders[0].LineStyle); // ExSkip

            foreach (Border border in borders) border.ClearFormatting();

            builder.CurrentParagraph.Runs[0].Text = "Paragraph with no border";

            doc.Save(ArtifactsDir + "Border.ClearFormatting.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Border.ClearFormatting.docx");

            foreach (Border testBorder in doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Borders)
            {
                Assert.AreEqual(Color.Empty.ToArgb(), testBorder.Color.ToArgb());
                Assert.AreEqual(0.0d, testBorder.LineWidth);
                Assert.AreEqual(LineStyle.None, testBorder.LineStyle);
            }
        }

        [Test]
        public void EqualityCountingAndVisibility()
        {
            //ExStart
            //ExFor:Border.Equals(Object)
            //ExFor:Border.Equals(Border)
            //ExFor:Border.GetHashCode
            //ExFor:Border.IsVisible
            //ExFor:BorderCollection.Count
            //ExFor:BorderCollection.Equals(BorderCollection)
            //ExFor:BorderCollection.Item(Int32)
            //ExSummary:Shows the equality of BorderCollections as well counting, visibility of their elements.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.CurrentParagraph.AppendChild(new Run(doc, "Paragraph 1."));

            Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
            BorderCollection firstParaBorders = firstParagraph.ParagraphFormat.Borders;

            builder.InsertParagraph();
            builder.CurrentParagraph.AppendChild(new Run(doc, "Paragraph 2."));

            Paragraph secondParagraph = builder.CurrentParagraph;
            BorderCollection secondParaBorders = secondParagraph.ParagraphFormat.Borders;

            // Two paragraphs have two different BorderCollections, but share the elements that are in from the first paragraph
            for (int i = 0; i < firstParaBorders.Count; i++)
            {
                Assert.IsTrue(firstParaBorders[i].Equals(secondParaBorders[i]));
                Assert.AreEqual(firstParaBorders[i].GetHashCode(), secondParaBorders[i].GetHashCode());

                // Borders are invisible by default
                Assert.IsFalse(firstParaBorders[i].IsVisible);
            }

            // Each border in the second paragraph collection becomes no longer the same as its counterpart from the first paragraph collection
            // Change all the elements in the second collection to make it completely different from the first
            Assert.AreEqual(6, secondParaBorders.Count); // ExSkip
            foreach (Border border in secondParaBorders)
            {
                border.LineStyle = LineStyle.DotDash;
            }

            // Now the BorderCollections both have their own elements
            for (int i = 0; i < firstParaBorders.Count; i++)
            {
                Assert.IsFalse(firstParaBorders[i].Equals(secondParaBorders[i]));
                Assert.AreNotEqual(firstParaBorders[i].GetHashCode(), secondParaBorders[i].GetHashCode());
                // Changing the line style made the borders visible
                Assert.IsTrue(secondParaBorders[i].IsVisible);
            }

            doc.Save(ArtifactsDir + "Border.EqualityCountingAndVisibility.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Border.EqualityCountingAndVisibility.docx");
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            foreach (Border testBorder in paragraphs[0].ParagraphFormat.Borders)
                Assert.AreEqual(LineStyle.None, testBorder.LineStyle);

            foreach (Border testBorder in paragraphs[1].ParagraphFormat.Borders)
                Assert.AreEqual(LineStyle.DotDash, testBorder.LineStyle);
        }

        [Test]
        public void VerticalAndHorizontalBorders()
        {
            //ExStart
            //ExFor:BorderCollection.Horizontal
            //ExFor:BorderCollection.Vertical
            //ExFor:Cell.LastParagraph
            //ExSummary:Shows the difference between the Horizontal and Vertical properties of BorderCollection.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // A BorderCollection is one of a Paragraph's formatting properties
            Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
            BorderCollection paragraphBorders = paragraph.ParagraphFormat.Borders;

            // paragraphBorders belongs to the first paragraph, but these changes will apply to subsequently created paragraphs
            paragraphBorders.Horizontal.Color = Color.Red;
            paragraphBorders.Horizontal.LineStyle = LineStyle.DashSmallGap;
            paragraphBorders.Horizontal.LineWidth = 3;

            // Horizontal borders only appear under a paragraph if there's another paragraph under it
            // Right now the first paragraph has no borders
            builder.CurrentParagraph.AppendChild(new Run(doc, "Paragraph above horizontal border."));

            // Now the first paragraph will have a red dashed line border under it
            // This new second paragraph can have a border too, but only if we add another paragraph underneath it
            builder.InsertParagraph();
            builder.CurrentParagraph.AppendChild(new Run(doc, "Paragraph below horizontal border."));

            // A table makes use of both vertical and horizontal properties of BorderCollection
            // Both these properties can only affect the inner borders of a table
            Table table = new Table(doc);
            doc.FirstSection.Body.AppendChild(table);

            for (int i = 0; i < 3; i++)
            {
                Row row = new Row(doc);
                BorderCollection rowBorders = row.RowFormat.Borders;

                // Vertical borders are ones between rows in a table
                rowBorders.Horizontal.Color = Color.Red;
                rowBorders.Horizontal.LineStyle = LineStyle.Dot;
                rowBorders.Horizontal.LineWidth = 2.0d;

                // Vertical borders are ones between cells in a table
                rowBorders.Vertical.Color = Color.Blue;
                rowBorders.Vertical.LineStyle = LineStyle.Dot;
                rowBorders.Vertical.LineWidth = 2.0d;

                // A blue dotted vertical border will appear between cells
                // A red dotted border will appear between rows
                row.AppendChild(new Cell(doc));
                row.LastCell.AppendChild(new Paragraph(doc));
                row.LastCell.FirstParagraph.AppendChild(new Run(doc, "Vertical border to the right."));

                row.AppendChild(new Cell(doc));
                row.LastCell.AppendChild(new Paragraph(doc));
                row.LastCell.LastParagraph.AppendChild(new Run(doc, "Vertical border to the left."));
                table.AppendChild(row);
            }

            doc.Save(ArtifactsDir + "Border.VerticalAndHorizontalBorders.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Border.VerticalAndHorizontalBorders.docx");
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            Assert.AreEqual(LineStyle.DashSmallGap, paragraphs[0].ParagraphFormat.Borders[BorderType.Horizontal].LineStyle);
            Assert.AreEqual(LineStyle.DashSmallGap, paragraphs[1].ParagraphFormat.Borders[BorderType.Horizontal].LineStyle);

            Table outTable = (Table)doc.GetChild(NodeType.Table, 0, true);

            foreach (Row row in outTable.GetChildNodes(NodeType.Row, true))
            {
                Assert.AreEqual(Color.Red.ToArgb(), row.RowFormat.Borders.Horizontal.Color.ToArgb());
                Assert.AreEqual(LineStyle.Dot, row.RowFormat.Borders.Horizontal.LineStyle);
                Assert.AreEqual(2.0d, row.RowFormat.Borders.Horizontal.LineWidth);

                Assert.AreEqual(Color.Blue.ToArgb(), row.RowFormat.Borders.Vertical.Color.ToArgb());
                Assert.AreEqual(LineStyle.Dot, row.RowFormat.Borders.Vertical.LineStyle);
                Assert.AreEqual(2.0d, row.RowFormat.Borders.Vertical.LineWidth);
            }
        }
    }
}