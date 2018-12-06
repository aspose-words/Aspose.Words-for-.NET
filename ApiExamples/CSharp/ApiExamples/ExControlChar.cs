using System;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExControlChar : ApiExampleBase
    {
        [Test]
        public void InsertControlChars()
        {
            //ExStart
            //ExFor:ControlChar.Cell
            //ExFor:ControlChar.ColumnBreak
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
            //ExSummary:Shows how to use various control characters.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a regular space
            builder.Write("Before space." + ControlChar.SpaceChar + "After space.");

            // Add a NBSP, or non-breaking space
            // Unlike the regular space, this space can't have an automatic line break at its position 
            builder.Write("Before space." + ControlChar.NonBreakingSpace + "After space.");

            // Add a tab character
            builder.Write("Before tab." + ControlChar.Tab + "After tab.");

            // Add a line break
            builder.Write("Before line break." + ControlChar.LineBreak + "After line break.");

            // This adds a new line and starts a new paragraph
            // Same value as ControlChar.Lf
            Assert.AreEqual(1, doc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, true).Count);
            builder.Write("Before line feed." + ControlChar.LineFeed + "After line feed.");
            Assert.AreEqual(2, doc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, true).Count);

            // The line feed character has two versions
            Assert.AreEqual(ControlChar.LineFeed, ControlChar.Lf);

            // Add a paragraph break, also adding a new paragraph
            builder.Write("Before paragraph break." + ControlChar.ParagraphBreak + "After paragraph break.");
            Assert.AreEqual(3, doc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, true).Count);

            // Add a section break. Note that this does not make a new section or paragraph
            Assert.AreEqual(1, doc.Sections.Count);
            builder.Write("Before section break." + ControlChar.SectionBreak + "After section break.");
            Assert.AreEqual(1, doc.Sections.Count);

            // A page break is the same value as a section break
            builder.Write("Before page break." + ControlChar.PageBreak + "After page break.");

            // We can add a new section like this
            doc.AppendChild(new Section(doc));
            builder.MoveToSection(1);

            // If you have a section with more than one column, you can use a column break to make following text start on a new column
            builder.CurrentSection.PageSetup.TextColumns.SetCount(2);
            builder.Write("Text at end of column 1." + ControlChar.ColumnBreak + "Text at beginning of column 2.");

            // Save document to see the characters we added
            doc.Save(ArtifactsDir + "ControlChar.Misc.docx");

            // There are char and string counterparts for most characters
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