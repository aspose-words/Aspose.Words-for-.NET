// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Collections.Generic;
using System.Drawing;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExBorderCollection : ApiExampleBase
    {
        [Test]
        public void GetBordersEnumerator()
        {
            //ExStart
            //ExFor:BorderCollection.GetEnumerator
            //ExSummary:Shows how to iterate over and edit all of the borders in a paragraph format object.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Configure the builder's paragraph format settings to create a green wave border on all sides.
            BorderCollection borders = builder.ParagraphFormat.Borders;

            using (IEnumerator<Border> enumerator = borders.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    Border border = enumerator.Current;
                    border.Color = Color.Green;
                    border.LineStyle = LineStyle.Wave;
                    border.LineWidth = 3;
                }
            }

            // Insert a paragraph. Our border settings will determine the appearance of its border.
            builder.Writeln("Hello world!");

            doc.Save(ArtifactsDir + "BorderCollection.GetBordersEnumerator.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "BorderCollection.GetBordersEnumerator.docx");

            foreach (Border border in doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Borders)
            {
                Assert.That(border.Color.ToArgb(), Is.EqualTo(Color.Green.ToArgb()));
                Assert.That(border.LineStyle, Is.EqualTo(LineStyle.Wave));
                Assert.That(border.LineWidth, Is.EqualTo(3.0d));
            }
        }

        [Test]
        public void RemoveAllBorders()
        {
            //ExStart
            //ExFor:BorderCollection.ClearFormatting
            //ExSummary:Shows how to remove all borders from all paragraphs in a document.
            Document doc = new Document(MyDir + "Borders.docx");

            // The first paragraph of this document has visible borders with these settings.
            BorderCollection firstParagraphBorders = doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Borders;

            Assert.That(firstParagraphBorders.Color.ToArgb(), Is.EqualTo(Color.Red.ToArgb()));
            Assert.That(firstParagraphBorders.LineStyle, Is.EqualTo(LineStyle.Single));
            Assert.That(firstParagraphBorders.LineWidth, Is.EqualTo(3.0d));

            // Use the "ClearFormatting" method on each paragraph to remove all borders.
            foreach (Paragraph paragraph in doc.FirstSection.Body.Paragraphs)
            {
                paragraph.ParagraphFormat.Borders.ClearFormatting();

                foreach (Border border in paragraph.ParagraphFormat.Borders)
                {
                    Assert.That(border.Color.ToArgb(), Is.EqualTo(Color.Empty.ToArgb()));
                    Assert.That(border.LineStyle, Is.EqualTo(LineStyle.None));
                    Assert.That(border.LineWidth, Is.EqualTo(0.0d));
                }
            }
            
            doc.Save(ArtifactsDir + "BorderCollection.RemoveAllBorders.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "BorderCollection.RemoveAllBorders.docx");

            foreach (Border border in doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Borders)
            {
                Assert.That(border.Color.ToArgb(), Is.EqualTo(Color.Empty.ToArgb()));
                Assert.That(border.LineStyle, Is.EqualTo(LineStyle.None));
                Assert.That(border.LineWidth, Is.EqualTo(0.0d));
            }
        }
    }
}