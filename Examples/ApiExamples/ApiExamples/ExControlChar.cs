// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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
            Assert.That(doc.GetText(), Is.EqualTo($"Hello world!{ControlChar.Cr}" +
                            $"Hello again!{ControlChar.Cr}" +
                            ControlChar.PageBreak));

            // When converting a document to string form,
            // we can omit some of the control characters with the Trim method.
            Assert.That(doc.GetText().Trim(), Is.EqualTo($"Hello world!{ControlChar.Cr}" +
                            "Hello again!"));
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
            Assert.That(doc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, true).Count, Is.EqualTo(1));
            builder.Write("Before line feed." + ControlChar.LineFeed + "After line feed.");
            Assert.That(doc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, true).Count, Is.EqualTo(2));

            // The line feed character has two versions.
            Assert.That(ControlChar.Lf, Is.EqualTo(ControlChar.LineFeed));

            // Carriage returns and line feeds can be represented together by one character.
            Assert.That(ControlChar.Cr + ControlChar.Lf, Is.EqualTo(ControlChar.CrLf));

            // Add a paragraph break, which will start a new paragraph.
            builder.Write("Before paragraph break." + ControlChar.ParagraphBreak + "After paragraph break.");
            Assert.That(doc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, true).Count, Is.EqualTo(3));

            // Add a section break. This does not make a new section or paragraph.
            Assert.That(doc.Sections.Count, Is.EqualTo(1));
            builder.Write("Before section break." + ControlChar.SectionBreak + "After section break.");
            Assert.That(doc.Sections.Count, Is.EqualTo(1));

            // Add a page break.
            builder.Write("Before page break." + ControlChar.PageBreak + "After page break.");

            // A page break is the same value as a section break.
            Assert.That(ControlChar.SectionBreak, Is.EqualTo(ControlChar.PageBreak));

            // Insert a new section, and then set its column count to two.
            doc.AppendChild(new Section(doc));
            builder.MoveToSection(1);
            builder.CurrentSection.PageSetup.TextColumns.SetCount(2);

            // We can use a control character to mark the point where text moves to the next column.
            builder.Write("Text at end of column 1." + ControlChar.ColumnBreak + "Text at beginning of column 2.");

            doc.Save(ArtifactsDir + "ControlChar.InsertControlChars.docx");

            // There are char and string counterparts for most characters.
            Assert.That(ControlChar.CellChar, Is.EqualTo(Convert.ToChar(ControlChar.Cell)));
            Assert.That(ControlChar.NonBreakingSpaceChar, Is.EqualTo(Convert.ToChar(ControlChar.NonBreakingSpace)));
            Assert.That(ControlChar.TabChar, Is.EqualTo(Convert.ToChar(ControlChar.Tab)));
            Assert.That(ControlChar.LineBreakChar, Is.EqualTo(Convert.ToChar(ControlChar.LineBreak)));
            Assert.That(ControlChar.LineFeedChar, Is.EqualTo(Convert.ToChar(ControlChar.LineFeed)));
            Assert.That(ControlChar.ParagraphBreakChar, Is.EqualTo(Convert.ToChar(ControlChar.ParagraphBreak)));
            Assert.That(ControlChar.SectionBreakChar, Is.EqualTo(Convert.ToChar(ControlChar.SectionBreak)));
            Assert.That(ControlChar.SectionBreakChar, Is.EqualTo(Convert.ToChar(ControlChar.PageBreak)));
            Assert.That(ControlChar.ColumnBreakChar, Is.EqualTo(Convert.ToChar(ControlChar.ColumnBreak)));
            //ExEnd
        }
    }
}