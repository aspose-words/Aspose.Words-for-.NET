// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.Border
{
    [TestFixture]
    public class ExBorder : QaTestsBase
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

            builder.Font.Border.Color = System.Drawing.Color.Green;
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

            Aspose.Words.Border topBorder = builder.ParagraphFormat.Borders[BorderType.Top];
            topBorder.Color = System.Drawing.Color.Red;
            topBorder.LineStyle = LineStyle.DashSmallGap;
            topBorder.LineWidth = 4;

            builder.Writeln("Hello World!");
            //ExEnd
        }

        [Test]
        public void ClearFormattingEx()
        {
            //ExStart
            //ExFor:ClearFormatting
            //ExId:ClearFormattingEx
            //ExSummary:Shows how to use ClearFormatting.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
            builder.Font.Border.ClearFormatting();
            //ExEnd
        }

        [Test]
        public void EqualsEx()
        {
            //ExStart
            //ExFor:Equals
            //ExId:EqualsEx
            //ExSummary:Shows how to use Equals.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
            Aspose.Words.Border border1 = builder.Font.Border;
            Aspose.Words.Border border2 = builder.Font.Border;

            Console.WriteLine(border1.Equals(border2));
            //ExEnd
        }

        [Test]
        public void GetHashCodeEx()
        {
            //ExStart
            //ExFor:GetHashCode
            //ExId:GetHashCodeEx
            //ExSummary:Shows how to use GetHashCode.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
            Aspose.Words.Border border = builder.Font.Border;

            int hash = border.GetHashCode();
            //ExEnd
        }
    }
}
