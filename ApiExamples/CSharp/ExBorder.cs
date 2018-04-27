// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
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
            //ExFor:DocumentBuilder.Write
            //ExSummary:Inserts a String surrounded by a border into a document.
            DocumentBuilder builder = new DocumentBuilder();

            builder.Font.Border.Color = Color.Green;
            builder.Font.Border.LineWidth = 2.5;
            builder.Font.Border.LineStyle = LineStyle.DashDotStroker;

            builder.Write("run of text in a green border");
            //ExEnd
        }

        [Test]
        public void ParagraphTopBorder()
        {
            //ExStart
            //ExFor:BorderCollection
            //ExFor:Border
            //ExFor:BorderType
            //ExFor:ParagraphFormat.Borders
            //ExSummary:Inserts a paragraph with a top border.
            DocumentBuilder builder = new DocumentBuilder();

            Border topBorder = builder.ParagraphFormat.Borders[BorderType.Top];
            topBorder.Color = Color.Red;
            topBorder.LineStyle = LineStyle.DashSmallGap;
            topBorder.LineWidth = 4;

            builder.Writeln("Hello World!");
            //ExEnd
        }

        [Test]
        public void ClearFormatting()
        {
            //ExStart
            //ExFor:Border.ClearFormatting
            //ExSummary:Shows how to remove borders from a paragraph one by one.
            Document doc = new Document(MyDir + "Border.Borders.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            BorderCollection borders = builder.ParagraphFormat.Borders;

            foreach (Border border in borders)
            {
                border.ClearFormatting();
            }

            builder.CurrentParagraph.Runs[0].Text = "Paragraph with no border";

            doc.Save(MyDir + @"\Artifacts\Border.NoBorder.doc");
            //ExEnd
        }

        [Test]
        public void Borders() // INSP: Please add a more informational name
        {
            //ExStart // INSP: You should use ExStart before ExFor tag
            //ExFor:Border.Equals(System.Object) // INSP: I can't see where you use this method. If you don't use, please add it or create new example.
            //ExFor:Border.GetHashCode
            //ExFor:Border.IsVisible
            //ExFor:BorderCollection.Count
            //ExFor:BorderCollection.Equals(BorderCollection)
            //ExFor:BorderCollection.Item(System.Int32)
            //ExSummary:Shows the equality of BorderCollections as well counting, visibility of their elements.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.CurrentParagraph.AppendChild(new Run(doc, "Paragraph 1."));

            Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
            BorderCollection firstParaBorders = firstParagraph.ParagraphFormat.Borders;

            // Borders are invisible by default
            foreach (Border border in firstParaBorders)
            {
                Assert.IsFalse(border.IsVisible);
            }

            // Changes to these borders in this paragraph will apply to subsequent paragraphs
            firstParaBorders[BorderType.Left].LineStyle = LineStyle.Double;
            firstParaBorders[BorderType.Right].LineStyle = LineStyle.Double;
            firstParaBorders[BorderType.Top].LineStyle = LineStyle.Double;
            firstParaBorders[BorderType.Bottom].LineStyle = LineStyle.Double;

            builder.InsertParagraph();
            builder.CurrentParagraph.AppendChild(new Run(doc, "Paragraph 2."));

            Paragraph secondParagraph = builder.CurrentParagraph;
            BorderCollection secondParaBorders = secondParagraph.ParagraphFormat.Borders;

            // Two paragraphs have two different BorderCollections, but share the elements from the first are given to the second
            for (int i = 0; i < firstParaBorders.Count; i++)
            {
                // INSP: We need to use the same code in example, don't need to add more and more functions. If you work only with LineStyle, so let's validate only LineStyle results.
                Assert.AreEqual(firstParaBorders[i].LineStyle, secondParaBorders[i].LineStyle);
                Assert.AreEqual(firstParaBorders[i].GetHashCode(), secondParaBorders[i].GetHashCode());
            }
            // INSP: We need to add more information about this line. Without information it seems unnecessary check.
            Assert.IsFalse(firstParaBorders.Equals(secondParaBorders));

            // If one BorderCollection element is changed in a subsequent paragraph, the rest must be changed too
            secondParaBorders[BorderType.Left].LineStyle = LineStyle.DotDash;
            secondParaBorders[BorderType.Right].LineStyle = LineStyle.DotDash;
            secondParaBorders[BorderType.Top].LineStyle = LineStyle.DotDash;
            secondParaBorders[BorderType.Bottom].LineStyle = LineStyle.DotDash;
            // INSP: If you use i < firstParaBorders.Count below, why we need to change secondParaBorders[BorderType.Vertical] and secondParaBorders[BorderType.Horizontal] ?
            secondParaBorders[BorderType.Vertical].LineStyle = LineStyle.DotDash;
            secondParaBorders[BorderType.Horizontal].LineStyle = LineStyle.DotDash;

            // Now the BorderCollections both have their own elements
            for (int i = 0; i < firstParaBorders.Count; i++)
            {
                Assert.AreNotEqual(firstParaBorders[i].LineStyle, secondParaBorders[i].LineStyle);
                Assert.AreNotEqual(firstParaBorders[i].GetHashCode(), secondParaBorders[i].GetHashCode());
            }
            //ExEnd
        }

        [Test]
        public void BordersVerticalAndHorizontal()
        {
            //ExStart
            //ExFor:BorderCollection.Horizontal
            //ExFor:BorderCollection.Vertical
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
                rowBorders.Horizontal.LineWidth = 2;

                // Vertical borders are ones between cells in a table
                rowBorders.Vertical.Color = Color.Blue;
                rowBorders.Vertical.LineStyle = LineStyle.Dot;
                rowBorders.Vertical.LineWidth = 2;

                // A blue dotted vertical border will appear between cells
                // A red dotted border will appear between rows
                row.AppendChild(new Cell(doc));
                row.LastCell.AppendChild(new Paragraph(doc));
                row.LastCell.FirstParagraph.AppendChild(new Run(doc, "Vertical border to the right."));

                row.AppendChild(new Cell(doc));
                row.LastCell.AppendChild(new Paragraph(doc));
                row.LastCell.FirstParagraph.AppendChild(new Run(doc, "Vertical border to the left."));
                table.AppendChild(row);
            }

            doc.Save(MyDir + @"\Artifacts\Border.HorizontalAndVerticalBorders.docx");
            //ExEnd
        }
    }
}
