// Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Themes;
using Xunit;
using Runner.MAUI;

namespace ApiExamples
{
    public class ExBorder : Dirs, IClassFixture<ConfigureTestsFixture>
    {
        [Fact]
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

            builder.Font.Border.Color = System.Drawing.Color.Green;
            builder.Font.Border.LineWidth = 2.5d;
            builder.Font.Border.LineStyle = LineStyle.DashDotStroker;

            builder.Write("Text surrounded by green border.");

            doc.Save(ArtifactsDir + "Border.FontBorder.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Border.FontBorder.docx");
            Aspose.Words.Border border = doc.FirstSection.Body.FirstParagraph.Runs[0].Font.Border;

            Assert.Equal(System.Drawing.Color.Green.ToArgb(), border.Color.ToArgb());
            Assert.Equal(2.5d, border.LineWidth);
            Assert.Equal(LineStyle.DashDotStroker, border.LineStyle);
        }

        [Fact]
        public void ParagraphTopBorder()
        {
            //ExStart
            //ExFor:BorderCollection
            //ExFor:Border.ThemeColor
            //ExFor:Border.TintAndShade
            //ExFor:Border
            //ExFor:BorderType
            //ExFor:ParagraphFormat.Borders
            //ExSummary:Shows how to insert a paragraph with a top border.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Aspose.Words.Border topBorder = builder.ParagraphFormat.Borders.Top;
            topBorder.LineWidth = 4.0d;
            topBorder.LineStyle = LineStyle.DashSmallGap;
            // Set ThemeColor only when LineWidth or LineStyle setted.
            topBorder.ThemeColor = ThemeColor.Accent1;
            topBorder.TintAndShade = 0.25d;

            builder.Writeln("Text with a top border.");

            doc.Save(ArtifactsDir + "Border.ParagraphTopBorder.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Border.ParagraphTopBorder.docx");
            Aspose.Words.Border border = doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Borders.Top;

            Assert.Equal(4.0d, border.LineWidth);
            Assert.Equal(LineStyle.DashSmallGap, border.LineStyle);
            Assert.Equal(ThemeColor.Accent1, border.ThemeColor);
            Assert.Equal(0.25d, border.TintAndShade, 1);
        }

        [Fact]
        public void ClearFormatting()
        {
            //ExStart
            //ExFor:Border.ClearFormatting
            //ExFor:Border.IsVisible
            //ExSummary:Shows how to remove borders from a paragraph.
            Document doc = new Document(MyDir + "Borders.docx");

            // Each paragraph has an individual set of borders.
            // We can access the settings for the appearance of these borders via the paragraph format object.
            BorderCollection borders = doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Borders;

            Assert.Equal(System.Drawing.Color.Red.ToArgb(), borders[0].Color.ToArgb());
            Assert.Equal(3.0d, borders[0].LineWidth);
            Assert.Equal(LineStyle.Single, borders[0].LineStyle);
            Assert.True(borders[0].IsVisible);

            // We can remove a border at once by running the ClearFormatting method. 
            // Running this method on every border of a paragraph will remove all its borders.
            foreach (Aspose.Words.Border border in borders)
                border.ClearFormatting();

            Assert.Equal(System.Drawing.Color.Empty.ToArgb(), borders[0].Color.ToArgb());
            Assert.Equal(0.0d, borders[0].LineWidth);
            Assert.Equal(LineStyle.None, borders[0].LineStyle);
            Assert.False(borders[0].IsVisible);

            doc.Save(ArtifactsDir + "Border.ClearFormatting.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Border.ClearFormatting.docx");

            foreach (Aspose.Words.Border testBorder in doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Borders)
            {
                Assert.Equal(System.Drawing.Color.Empty.ToArgb(), testBorder.Color.ToArgb());
                Assert.Equal(0.0d, testBorder.LineWidth);
                Assert.Equal(LineStyle.None, testBorder.LineStyle);
            }
        }

        [Fact]
        public void SharedElements()
        {
            //ExStart
            //ExFor:Border.Equals(Object)
            //ExFor:Border.Equals(Border)
            //ExFor:Border.GetHashCode
            //ExFor:BorderCollection.Count
            //ExFor:BorderCollection.Equals(BorderCollection)
            //ExFor:BorderCollection.Item(Int32)
            //ExSummary:Shows how border collections can share elements.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Paragraph 1.");
            builder.Write("Paragraph 2.");

            // Since we used the same border configuration while creating
            // these paragraphs, their border collections share the same elements.
            BorderCollection firstParagraphBorders = doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Borders;
            BorderCollection secondParagraphBorders = builder.CurrentParagraph.ParagraphFormat.Borders;
            Assert.Equal(6, firstParagraphBorders.Count); //ExSkip

            for (int i = 0; i < firstParagraphBorders.Count; i++)
            {
                Assert.True(firstParagraphBorders[i].Equals(secondParagraphBorders[i]));
                Assert.Equal(firstParagraphBorders[i].GetHashCode(), secondParagraphBorders[i].GetHashCode());
                Assert.False(firstParagraphBorders[i].IsVisible);
            }

            foreach (Aspose.Words.Border border in secondParagraphBorders)
                border.LineStyle = LineStyle.DotDash;

            // After changing the line style of the borders in just the second paragraph,
            // the border collections no longer share the same elements.
            for (int i = 0; i < firstParagraphBorders.Count; i++)
            {
                Assert.False(firstParagraphBorders[i].Equals(secondParagraphBorders[i]));
                Assert.NotEqual(firstParagraphBorders[i].GetHashCode(), secondParagraphBorders[i].GetHashCode());

                // Changing the appearance of an empty border makes it visible.
                Assert.True(secondParagraphBorders[i].IsVisible);
            }

            doc.Save(ArtifactsDir + "Border.SharedElements.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Border.SharedElements.docx");
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            foreach (Aspose.Words.Border testBorder in paragraphs[0].ParagraphFormat.Borders)
                Assert.Equal(LineStyle.None, testBorder.LineStyle);

            foreach (Aspose.Words.Border testBorder in paragraphs[1].ParagraphFormat.Borders)
                Assert.Equal(LineStyle.DotDash, testBorder.LineStyle);
        }

        [Fact]
        public void HorizontalBorders()
        {
            //ExStart
            //ExFor:BorderCollection.Horizontal
            //ExSummary:Shows how to apply settings to horizontal borders to a paragraph's format.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a red horizontal border for the paragraph. Any paragraphs created afterwards will inherit these border settings.
            BorderCollection borders = doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Borders;
            borders.Horizontal.Color = System.Drawing.Color.Red;
            borders.Horizontal.LineStyle = LineStyle.DashSmallGap;
            borders.Horizontal.LineWidth = 3;

            // Write text to the document without creating a new paragraph afterward.
            // Since there is no paragraph underneath, the horizontal border will not be visible.
            builder.Write("Paragraph above horizontal border.");

            // Once we add a second paragraph, the border of the first paragraph will become visible.
            builder.InsertParagraph();
            builder.Write("Paragraph below horizontal border.");

            doc.Save(ArtifactsDir + "Border.HorizontalBorders.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Border.HorizontalBorders.docx");
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            Assert.Equal(LineStyle.DashSmallGap, paragraphs[0].ParagraphFormat.Borders[BorderType.Horizontal].LineStyle);
            Assert.Equal(LineStyle.DashSmallGap, paragraphs[1].ParagraphFormat.Borders[BorderType.Horizontal].LineStyle);
        }

        [Fact]
        public void VerticalBorders()
        {
            //ExStart
            //ExFor:BorderCollection.Horizontal
            //ExFor:BorderCollection.Vertical
            //ExFor:Cell.LastParagraph
            //ExSummary:Shows how to apply settings to vertical borders to a table row's format.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a table with red and blue inner borders.
            Table table = builder.StartTable();

            for (int i = 0; i < 3; i++)
            {
                builder.InsertCell();
                builder.Write($"Row {i + 1}, Column 1");
                builder.InsertCell();
                builder.Write($"Row {i + 1}, Column 2");

                Row row = builder.EndRow();
                BorderCollection borders = row.RowFormat.Borders;

                // Adjust the appearance of borders that will appear between rows.
                borders.Horizontal.Color = System.Drawing.Color.Red;
                borders.Horizontal.LineStyle = LineStyle.Dot;
                borders.Horizontal.LineWidth = 2.0d;

                // Adjust the appearance of borders that will appear between cells.
                borders.Vertical.Color = System.Drawing.Color.Blue;
                borders.Vertical.LineStyle = LineStyle.Dot;
                borders.Vertical.LineWidth = 2.0d;
            }

            // A row format, and a cell's inner paragraph use different border settings.
            Aspose.Words.Border border = table.FirstRow.FirstCell.LastParagraph.ParagraphFormat.Borders.Vertical;

            Assert.Equal(System.Drawing.Color.Empty.ToArgb(), border.Color.ToArgb());
            Assert.Equal(0.0d, border.LineWidth);
            Assert.Equal(LineStyle.None, border.LineStyle);

            doc.Save(ArtifactsDir + "Border.VerticalBorders.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Border.VerticalBorders.docx");
            table = doc.FirstSection.Body.Tables[0];

            foreach (Row row in table.GetChildNodes(NodeType.Row, true))
            {
                Assert.Equal(System.Drawing.Color.Red.ToArgb(), row.RowFormat.Borders.Horizontal.Color.ToArgb());
                Assert.Equal(LineStyle.Dot, row.RowFormat.Borders.Horizontal.LineStyle);
                Assert.Equal(2.0d, row.RowFormat.Borders.Horizontal.LineWidth);

                Assert.Equal(System.Drawing.Color.Blue.ToArgb(), row.RowFormat.Borders.Vertical.Color.ToArgb());
                Assert.Equal(LineStyle.Dot, row.RowFormat.Borders.Vertical.LineStyle);
                Assert.Equal(2.0d, row.RowFormat.Borders.Vertical.LineWidth);
            }
        }
    }
}