using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExControlChar : ApiExampleBase
    {
        [Test]
        public void ControlChar()
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
            DocumentBuilder db = new DocumentBuilder(doc);

            // Add a space.
            db.Writeln("Before space." + Aspose.Words.ControlChar.SpaceChar + "After space.");

            // Add a space.
            db.Writeln("Before space." + Aspose.Words.ControlChar.NonBreakingSpace + "After space.");

            // Add a tab character.
            db.Writeln("Before tab." + Aspose.Words.ControlChar.Tab + "After tab.");

            // These are all a new line. A new paragraph is also started following each character.
            db.Writeln("Before line break." + Aspose.Words.ControlChar.LineBreak + "After line break.");
            db.Writeln("Before line feed." + Aspose.Words.ControlChar.LineFeed + "After line feed.");
            db.Writeln("Before lf." + Aspose.Words.ControlChar.Lf + "After lf.");

            // Add a paragraph break, also adding a new paragraph.
            db.Writeln("Before paragraph break." + Aspose.Words.ControlChar.ParagraphBreak + "After paragraph break.");

            // Add a section break. Note that this does not make a new section.
            Assert.AreEqual(1, doc.Sections.Count);
            db.Writeln("Before section break." + Aspose.Words.ControlChar.SectionBreak + "After section break.");
            Assert.AreEqual(1, doc.Sections.Count);

            // Add a page break. Same value as a section break.
            db.Writeln("Before page break." + Aspose.Words.ControlChar.PageBreak + "After page break.");

            // so if we want a second section for this document, we have to make it manually.
            doc.AppendChild(new Section(doc));
            db.MoveToSection(1);

            // If you have a section with more than one column, you can force subsequent text onto the next column with a control character.
            db.CurrentSection.PageSetup.TextColumns.SetCount(2);
            db.Writeln("Text at end of column 1." + Aspose.Words.ControlChar.ColumnBreak + "Text at beginning of column 2.");

            doc.Save(MyDir + @"\Artifacts\ControlChar.doc");
            //ExEnd
        }
    }
}
