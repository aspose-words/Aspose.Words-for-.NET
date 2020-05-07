// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExStyles : ApiExampleBase
    {
        [Test]
        public void Styles()
        {
            //ExStart
            //ExFor:DocumentBase.Styles
            //ExFor:Style.Document
            //ExFor:Style.Name
            //ExFor:Style.IsHeading
            //ExFor:Style.IsQuickStyle
            //ExFor:Style.NextParagraphStyleName
            //ExFor:Style.Styles
            //ExFor:Style.Type
            //ExFor:StyleCollection.Document
            //ExFor:StyleCollection.GetEnumerator
            //ExSummary:Shows how to get access to the collection of styles defined in the document.
            Document doc = new Document();
           
            using (IEnumerator<Style> stylesEnum = doc.Styles.GetEnumerator())
            {
                while (stylesEnum.MoveNext())
                {
                    Style curStyle = stylesEnum.Current;
                    Console.WriteLine($"Style name:\t\"{curStyle.Name}\", of type \"{curStyle.Type}\"");
                    Console.WriteLine($"\tSubsequent style:\t{curStyle.NextParagraphStyleName}");
                    Console.WriteLine($"\tIs heading:\t\t\t{curStyle.IsHeading}");
                    Console.WriteLine($"\tIs QuickStyle:\t\t{curStyle.IsQuickStyle}");

                    Assert.AreEqual(doc, curStyle.Document);
                }
            }
            //ExEnd
        }

        [Test]
        public void StyleCollection()
        {
            //ExStart
            //ExFor:StyleCollection.Add(Style)
            //ExFor:StyleCollection.Count
            //ExFor:StyleCollection.DefaultFont
            //ExFor:StyleCollection.DefaultParagraphFormat
            //ExFor:StyleCollection.Item(StyleIdentifier)
            //ExFor:StyleCollection.Item(Int32)
            //ExSummary:Shows how to add a Style to a StyleCollection.
            Document doc = new Document();

            // New documents come with a collection of default styles that can be applied to paragraphs
            StyleCollection styles = doc.Styles;
            Assert.AreEqual(4, styles.Count);

            // We can set default parameters for new styles that will be added to the collection from now on
            styles.DefaultFont.Name = "Courier New";
            styles.DefaultParagraphFormat.FirstLineIndent = 15.0;

            styles.Add(StyleType.Paragraph, "MyStyle");

            // Styles within the collection can be referenced either by index or name
            Assert.AreEqual("Courier New", styles[4].Font.Name);
            Assert.AreEqual(15.0, styles["MyStyle"].ParagraphFormat.FirstLineIndent);
            //ExEnd
        }


        [Test]
        public void SetAllStyles()
        {
            //ExStart
            //ExFor:Style.Font
            //ExFor:Style
            //ExSummary:Shows how to change the font formatting of all styles in a document.
            Document doc = new Document();
            foreach (Style style in doc.Styles)
            {
                if (style.Font != null)
                {
                    style.Font.ClearFormatting();
                    style.Font.Size = 20;
                    style.Font.Name = "Arial";
                }
            }
            //ExEnd
        }

        [Test]
        public void ChangeStyleOfTocLevel()
        {
            Document doc = new Document();
            
            // Retrieve the style used for the first level of the TOC and change the formatting of the style
            doc.Styles[StyleIdentifier.Toc1].Font.Bold = true;
        }

        [Test]
        public void ChangeTocsTabStops()
        {
            //ExStart
            //ExFor:TabStop
            //ExFor:ParagraphFormat.TabStops
            //ExFor:Style.StyleIdentifier
            //ExFor:TabStopCollection.RemoveByPosition
            //ExFor:TabStop.Alignment
            //ExFor:TabStop.Position
            //ExFor:TabStop.Leader
            //ExSummary:Shows how to modify the position of the right tab stop in TOC related paragraphs.
            Document doc = new Document(MyDir + "Table of contents.docx");

            // Iterate through all paragraphs in the document
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true).OfType<Paragraph>())
            {
                // Check if this paragraph is formatted using the TOC result based styles. This is any style between TOC and TOC9
                if (para.ParagraphFormat.Style.StyleIdentifier >= StyleIdentifier.Toc1 &&
                    para.ParagraphFormat.Style.StyleIdentifier <= StyleIdentifier.Toc9)
                {
                    // Get the first tab used in this paragraph, this should be the tab used to align the page numbers
                    TabStop tab = para.ParagraphFormat.TabStops[0];
                    // Remove the old tab from the collection
                    para.ParagraphFormat.TabStops.RemoveByPosition(tab.Position);
                    // Insert a new tab using the same properties but at a modified position
                    // We could also change the separators used (dots) by passing a different Leader type
                    para.ParagraphFormat.TabStops.Add(tab.Position - 50, tab.Alignment, tab.Leader);
                }
            }

            doc.Save(ArtifactsDir + "Styles.ChangeTocsTabStops.docx");
            //ExEnd
        }

        [Test]
        public void CopyStyleSameDocument()
        {
            Document doc = new Document(MyDir + "Document.docx");

            //ExStart
            //ExFor:StyleCollection.AddCopy
            //ExFor:Style.Name
            //ExSummary:Demonstrates how to copy a style within the same document.
            // The AddCopy method creates a copy of the specified style and automatically generates a new name for the style, such as "Heading 1_0"
            Style newStyle = doc.Styles.AddCopy(doc.Styles["Heading 1"]);

            // You can change the new style name if required as the Style.Name property is read-write
            newStyle.Name = "My Heading 1";
            //ExEnd

            Assert.NotNull(newStyle);
            Assert.AreEqual("My Heading 1", newStyle.Name);
            Assert.AreEqual(doc.Styles["Heading 1"].Type, newStyle.Type);
        }

        [Test]
        public void CopyStyleDifferentDocument()
        {
            Document dstDoc = new Document();
            Document srcDoc = new Document();

            //ExStart
            //ExFor:StyleCollection.AddCopy
            //ExSummary:Demonstrates how to copy style from one document into a different document.
            // This is the style in the source document to copy to the destination document
            Style srcStyle = srcDoc.Styles[StyleIdentifier.Heading1];

            // Change the font of the heading style to red
            srcStyle.Font.Color = Color.Red;

            // The AddCopy method can be used to copy a style from a different document
            Style newStyle = dstDoc.Styles.AddCopy(srcStyle);
            //ExEnd

            Assert.NotNull(newStyle);
            Assert.AreEqual("Heading 1", newStyle.Name);
            Assert.AreEqual(Color.Red.ToArgb(), newStyle.Font.Color.ToArgb());
        }

        [Test]
        public void OverwriteStyleDifferentDocument()
        {
            Document dstDoc = new Document();
            Document srcDoc = new Document();

            //ExStart
            //ExFor:StyleCollection.AddCopy
            //ExSummary:Demonstrates how to copy a style from one document to another and override an existing style in the destination document.
            // This is the style in the source document to copy to the destination document
            Style srcStyle = srcDoc.Styles[StyleIdentifier.Heading1];

            // Change the font of the heading style to red
            srcStyle.Font.Color = Color.Red;

            // The AddCopy method can be used to copy a style to a different document
            Style newStyle = dstDoc.Styles.AddCopy(srcStyle);

            // The name of the new style can be changed to the name of any existing style. Doing this will override the existing style
            newStyle.Name = "Heading 1";
            //ExEnd

            Assert.NotNull(newStyle);
            Assert.AreEqual("Heading 1", newStyle.Name);
            Assert.IsNull(dstDoc.Styles["Heading 1_0"]);
            Assert.AreEqual(Color.Red.ToArgb(), newStyle.Font.Color.ToArgb());
        }

        [Test]
        public void DefaultStyles()
        {
            Document doc = new Document();

            // Add document-wide defaults parameters
            doc.Styles.DefaultFont.Name = "PMingLiU";
            doc.Styles.DefaultFont.Bold = true;

            doc.Styles.DefaultParagraphFormat.SpaceAfter = 20;
            doc.Styles.DefaultParagraphFormat.Alignment = ParagraphAlignment.Right;

            doc = DocumentHelper.SaveOpen(doc);

            Assert.IsTrue(doc.Styles.DefaultFont.Bold);
            Assert.AreEqual("PMingLiU", doc.Styles.DefaultFont.Name);
            Assert.AreEqual(20, doc.Styles.DefaultParagraphFormat.SpaceAfter);
            Assert.AreEqual(ParagraphAlignment.Right, doc.Styles.DefaultParagraphFormat.Alignment);
        }

        [Test]
        public void ParagraphStyleBulleted()
        {
            //ExStart
            //ExFor:StyleCollection
            //ExFor:DocumentBase.Styles
            //ExFor:Style
            //ExFor:Font
            //ExFor:Style.Font
            //ExFor:Style.ParagraphFormat
            //ExFor:Style.ListFormat
            //ExFor:ParagraphFormat.Style
            //ExSummary:Shows how to create and use a paragraph style with list formatting.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a paragraph style and specify some formatting for it
            Style style = doc.Styles.Add(StyleType.Paragraph, "MyStyle1");
            style.Font.Size = 24;
            style.Font.Name = "Verdana";
            style.ParagraphFormat.SpaceAfter = 12;

            // Create a list and make sure the paragraphs that use this style will use this list
            style.ListFormat.List = doc.Lists.Add(ListTemplate.BulletDefault);
            style.ListFormat.ListLevelNumber = 0;

            // Apply the paragraph style to the current paragraph in the document and add some text
            builder.ParagraphFormat.Style = style;
            builder.Writeln("Hello World: MyStyle1, bulleted.");

            // Change to a paragraph style that has no list formatting
            builder.ParagraphFormat.Style = doc.Styles["Normal"];
            builder.Writeln("Hello World: Normal.");

            builder.Document.Save(ArtifactsDir + "Styles.ParagraphStyleBulleted.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Styles.ParagraphStyleBulleted.docx");

            style = doc.Styles["MyStyle1"];

            Assert.AreEqual("MyStyle1", style.Name);
            Assert.AreEqual(24, style.Font.Size);
            Assert.AreEqual("Verdana", style.Font.Name);
            Assert.AreEqual(12.0d, style.ParagraphFormat.SpaceAfter);
        }

        [Test]
        public void RemoveEx()
        {
            //ExStart
            //ExFor:Style.Remove
            //ExSummary:Shows how to pick a style that is defined in the document and remove it.
            Document doc = new Document();
            doc.Styles["Normal"].Remove();
            //ExEnd
        }

        [Test]
        public void StyleAliases()
        {
            //ExStart
            //ExFor:Style.Aliases
            //ExFor:Style.BaseStyleName
            //ExFor:Style.Equals(Aspose.Words.Style)
            //ExFor:Style.LinkedStyleName
            //ExSummary:Shows how to use style aliases.
            // Open a document that had a style inserted with commas in its name which separate the style name and aliases
            Document doc = new Document(MyDir + "Style with alias.docx");

            // The aliases, separate from the name can be found here
            Style style = doc.Styles["MyStyle"];
            Assert.AreEqual(new [] { "MyStyle Alias 1", "MyStyle Alias 2" }, style.Aliases);
            Assert.AreEqual("Title", style.BaseStyleName);
            Assert.AreEqual("MyStyle Char", style.LinkedStyleName);

            // A style can be referenced by alias as well as name
            Assert.IsTrue(style.Equals(doc.Styles["MyStyle Alias 1"]));
            //ExEnd
        }
    }
}