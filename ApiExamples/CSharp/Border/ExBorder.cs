// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Drawing;
using Aspose.Words;
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
            //ExSummary:Inserts a string surrounded by a border into a document.
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
            //ExFor:DocumentBuilder.ParagraphFormat
            //ExFor:DocumentBuilder.Writeln(String)
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
        public void ClearFormattingEx()
        {
            //ExStart
            //ExFor:Border.ClearFormatting
            //ExSummary:Shows how to remove borders from a paragraph one by one.
            Document doc = new Document(MyDir + "Document.Borders.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);
            BorderCollection borders = builder.ParagraphFormat.Borders;

            foreach (Border border in borders)
            {
                border.ClearFormatting();
            }

            builder.CurrentParagraph.Runs[0].Text = "Paragraph with no border";
            doc.Save(MyDir + @"\Artifacts\Document.NoBorder.doc");
            //ExEnd
        }
    }
}