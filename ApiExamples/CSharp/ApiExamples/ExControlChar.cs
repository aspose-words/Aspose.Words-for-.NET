// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExControlChar : ApiExampleBase
    {
        [Test]
        public void CarriageReturn()
        {
            //ExStart
            //ExFor:ControlChar
            //ExFor:ControlChar.Cr
            //ExFor:Node.GetText
            //ExSummary:Shows how to use control characters.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert paragraphs with text with DocumentBuilder.
            builder.Writeln("Hello world!");
            builder.Writeln("Hello again!");

            // Converting the document to text form reveals that control characters
            // represent some of the document's structural elements, such as page breaks.
            Assert.AreEqual($"Hello world!{ControlChar.Cr}" +
                            $"Hello again!{ControlChar.Cr}" +
                            ControlChar.PageBreak, doc.GetText());

            // When converting a document to string form,
            // we can omit some of the control characters with the Trim method.
            Assert.AreEqual($"Hello world!{ControlChar.Cr}" +
                            "Hello again!", doc.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void InsertControlChars()
        {
            //ExStart
            //ExFor:ControlChar.Cell
            //ExFor:ControlChar.ColumnBreak
            //ExFor:ControlChar.CrLf
            //ExFor:ControlChar.Lf
            //ExFor:ControlChar.LineBreak
            //ExFor:ControlChar.LineFeed
            //ExFor:ControlChar.NonBreakingSpace
            //ExFor:ControlChar.PageBreak
            //ExFor:ControlChar.ParagraphBreak
            //ExFor:ControlChar.SectionBreak
            //ExFor:ControlChar.CellChar
            //ExFor:ControlChar.ColumnBreakChar
            //ExFor:ControlChar.DefaultTextInputChar
            //ExFor:ControlChar.FieldEndChar
            //ExFor:ControlChar.FieldStartChar
            //ExFor:ControlChar.FieldSeparatorChar
            //ExFor:ControlChar.LineBreakChar
            //ExFor:ControlChar.LineFeedChar
            //ExFor:ControlChar.NonBreakingHyphenChar
            //ExFor:ControlChar.NonBreakingSpaceChar
            //ExFor:ControlChar.OptionalHyphenChar
            //ExFor:ControlChar.PageBreakChar
            //ExFor:ControlChar.ParagraphBreakChar
            //ExFor:ControlChar.SectionBreakChar
            //ExFor:ControlChar.SpaceChar
            //ExSummary:Shows how to add various control characters to a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a regular space.
            builder.Write("Before space." + ControlChar.SpaceChar + "After space.");

            // Add an NBSP, which is a non-breaking space.
            // Unlike the regular space, this space cannot have an automatic line break at its position.
            builder.Write("Before space." + ControlChar.NonBreakingSpace + "After space.");

            // Add a tab character.
            builder.Write("Before tab." + ControlChar.Tab + "After tab.");

            // Add a line break.
            builder.Write("Before line break." + ControlChar.LineBreak + "After line break.");

            // Add a new line and starts a new paragraph.
            Assert.AreEqual(1, doc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, true).Count);
            builder.Write("Before line feed." + ControlChar.LineFeed + "After line feed.");
            Assert.AreEqual(2, doc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, true).Count);

            // The line feed character has two versions.
            Assert.AreEqual(ControlChar.LineFeed, ControlChar.Lf);

            // Carriage returns and line feeds can be represented together by one character.
            Assert.AreEqual(ControlChar.CrLf, ControlChar.Cr + ControlChar.Lf);

            // Add a paragraph break, which will start a new paragraph.
            builder.Write("Before paragraph break." + ControlChar.ParagraphBreak + "After paragraph break.");
            Assert.AreEqual(3, doc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, true).Count);

            // Add a section break. This does not make a new section or paragraph.
            Assert.AreEqual(1, doc.Sections.Count);
            builder.Write("Before section break." + ControlChar.SectionBreak + "After section break.");
            Assert.AreEqual(1, doc.Sections.Count);

            // Add a page break.
            builder.Write("Before page break." + ControlChar.PageBreak + "After page break.");

            // A page break is the same value as a section break.
            Assert.AreEqual(ControlChar.PageBreak, ControlChar.SectionBreak);

            // Insert a new section, and then set its column count to two.
            doc.AppendChild(new Section(doc));
            builder.MoveToSection(1);
            builder.CurrentSection.PageSetup.TextColumns.SetCount(2);

            // We can use a control character to mark the point where text moves to the next column.
            builder.Write("Text at end of column 1." + ControlChar.ColumnBreak + "Text at beginning of column 2.");

            doc.Save(ArtifactsDir + "ControlChar.InsertControlChars.docx");

            // There are char and string counterparts for most characters.
            Assert.AreEqual(Convert.ToChar(ControlChar.Cell), ControlChar.CellChar);
            Assert.AreEqual(Convert.ToChar(ControlChar.NonBreakingSpace), ControlChar.NonBreakingSpaceChar);
            Assert.AreEqual(Convert.ToChar(ControlChar.Tab), ControlChar.TabChar);
            Assert.AreEqual(Convert.ToChar(ControlChar.LineBreak), ControlChar.LineBreakChar);
            Assert.AreEqual(Convert.ToChar(ControlChar.LineFeed), ControlChar.LineFeedChar);
            Assert.AreEqual(Convert.ToChar(ControlChar.ParagraphBreak), ControlChar.ParagraphBreakChar);
            Assert.AreEqual(Convert.ToChar(ControlChar.SectionBreak), ControlChar.SectionBreakChar);
            Assert.AreEqual(Convert.ToChar(ControlChar.PageBreak), ControlChar.SectionBreakChar);
            Assert.AreEqual(Convert.ToChar(ControlChar.ColumnBreak), ControlChar.ColumnBreakChar);
            //ExEnd
        }
    }
}