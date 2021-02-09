// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
            //ExSummary:Shows how to access a document's style collection.
            Document doc = new Document();
           
            Assert.AreEqual(4, doc.Styles.Count);
            
            // Enumerate and list all the styles that a document created using Aspose.Words contains by default.
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
        public void CreateStyle()
        {
            //ExStart
            //ExFor:Style.Font
            //ExFor:Style
            //ExFor:Style.Remove
            //ExSummary:Shows how to create and apply a custom style.
            Document doc = new Document();

            Style style = doc.Styles.Add(StyleType.Paragraph, "MyStyle");
            style.Font.Name = "Times New Roman";
            style.Font.Size = 16;
            style.Font.Color = Color.Navy;

            DocumentBuilder builder = new DocumentBuilder(doc);

            // Apply one of the styles from the document to the paragraph that the document builder is creating.
            builder.ParagraphFormat.Style = doc.Styles["MyStyle"];
            builder.Writeln("Hello world!");

            Style firstParagraphStyle = doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Style;

            Assert.AreEqual(style, firstParagraphStyle);

            // Remove our custom style from the document's styles collection.
            doc.Styles["MyStyle"].Remove();

            firstParagraphStyle = doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Style;

            // Any text that used a removed style reverts to the default formatting.
            Assert.False(doc.Styles.Any(s => s.Name == "MyStyle"));
            Assert.AreEqual("Times New Roman", firstParagraphStyle.Font.Name);
            Assert.AreEqual(12.0d, firstParagraphStyle.Font.Size);
            Assert.AreEqual(Color.Empty.ToArgb(), firstParagraphStyle.Font.Color.ToArgb());
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
            //ExSummary:Shows how to add a Style to a document's styles collection.
            Document doc = new Document();
            StyleCollection styles = doc.Styles;

            // Set default parameters for new styles that we may later add to this collection.
            styles.DefaultFont.Name = "Courier New";

            // If we add a style of the "StyleType.Paragraph", the collection will apply the values of
            // its "DefaultParagraphFormat" property to the style's "ParagraphFormat" property.
            styles.DefaultParagraphFormat.FirstLineIndent = 15.0;

            // Add a style, and then verify that it has the default settings.
            styles.Add(StyleType.Paragraph, "MyStyle");

            Assert.AreEqual("Courier New", styles[4].Font.Name);
            Assert.AreEqual(15.0, styles["MyStyle"].ParagraphFormat.FirstLineIndent);
            //ExEnd
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

            // Iterate through all paragraphs with TOC result-based styles; this is any style between TOC and TOC9.
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true).OfType<Paragraph>())
                if (para.ParagraphFormat.Style.StyleIdentifier >= StyleIdentifier.Toc1 &&
                    para.ParagraphFormat.Style.StyleIdentifier <= StyleIdentifier.Toc9)
                {
                    // Get the first tab used in this paragraph, this should be the tab used to align the page numbers.
                    TabStop tab = para.ParagraphFormat.TabStops[0];

                    // Replace the first default tab, stop with a custom tab stop.
                    para.ParagraphFormat.TabStops.RemoveByPosition(tab.Position);
                    para.ParagraphFormat.TabStops.Add(tab.Position - 50, tab.Alignment, tab.Leader);
                }

            doc.Save(ArtifactsDir + "Styles.ChangeTocsTabStops.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Styles.ChangeTocsTabStops.docx");

            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true).OfType<Paragraph>())
                if (para.ParagraphFormat.Style.StyleIdentifier >= StyleIdentifier.Toc1 &&
                    para.ParagraphFormat.Style.StyleIdentifier <= StyleIdentifier.Toc9)
                {
                    TabStop tabStop = para.GetEffectiveTabStops()[0];
                    Assert.AreEqual(400.8d, tabStop.Position);
                    Assert.AreEqual(TabAlignment.Right, tabStop.Alignment);
                    Assert.AreEqual(TabLeader.Dots, tabStop.Leader);
                }
        }

        [Test]
        public void CopyStyleSameDocument()
        {
            //ExStart
            //ExFor:StyleCollection.AddCopy
            //ExFor:Style.Name
            //ExSummary:Shows how to clone a document's style.
            Document doc = new Document();

            // The AddCopy method creates a copy of the specified style and
            // automatically generates a new name for the style, such as "Heading 1_0".
            Style newStyle = doc.Styles.AddCopy(doc.Styles["Heading 1"]);

            // Use the style's "Name" property to change the style's identifying name.
            newStyle.Name = "My Heading 1";

            // Our document now has two identical looking styles with different names.
            // Changing settings of one of the styles do not affect the other.
            newStyle.Font.Color = Color.Red;

            Assert.AreEqual("My Heading 1", newStyle.Name);
            Assert.AreEqual("Heading 1", doc.Styles["Heading 1"].Name);

            Assert.AreEqual(doc.Styles["Heading 1"].Type, newStyle.Type);
            Assert.AreEqual(doc.Styles["Heading 1"].Font.Name, newStyle.Font.Name);
            Assert.AreEqual(doc.Styles["Heading 1"].Font.Size, newStyle.Font.Size);
            Assert.AreNotEqual(doc.Styles["Heading 1"].Font.Color, newStyle.Font.Color);
            //ExEnd
        }

        [Test]
        public void CopyStyleDifferentDocument()
        {
            //ExStart
            //ExFor:StyleCollection.AddCopy
            //ExSummary:Shows how to import a style from one document into a different document.
            Document srcDoc = new Document();

            // Create a custom style for the source document.
            Style srcStyle = srcDoc.Styles.Add(StyleType.Paragraph, "MyStyle");
            srcStyle.Font.Color = Color.Red;

            // Import the source document's custom style into the destination document.
            Document dstDoc = new Document();
            Style newStyle = dstDoc.Styles.AddCopy(srcStyle);

            // The imported style has an appearance identical to its source style.
            Assert.AreEqual("MyStyle", newStyle.Name);
            Assert.AreEqual(Color.Red.ToArgb(), newStyle.Font.Color.ToArgb());
            //ExEnd
        }

        [Test]
        public void DefaultStyles()
        {
            Document doc = new Document();

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
        public void ParagraphStyleBulletedList()
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

            // Create a custom paragraph style.
            Style style = doc.Styles.Add(StyleType.Paragraph, "MyStyle1");
            style.Font.Size = 24;
            style.Font.Name = "Verdana";
            style.ParagraphFormat.SpaceAfter = 12;

            // Create a list and make sure the paragraphs that use this style will use this list.
            style.ListFormat.List = doc.Lists.Add(ListTemplate.BulletDefault);
            style.ListFormat.ListLevelNumber = 0;

            // Apply the paragraph style to the document builder's current paragraph, and then add some text.
            builder.ParagraphFormat.Style = style;
            builder.Writeln("Hello World: MyStyle1, bulleted list.");

            // Change the document builder's style to one that has no list formatting and write another paragraph.
            builder.ParagraphFormat.Style = doc.Styles["Normal"];
            builder.Writeln("Hello World: Normal.");

            builder.Document.Save(ArtifactsDir + "Styles.ParagraphStyleBulletedList.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Styles.ParagraphStyleBulletedList.docx");

            style = doc.Styles["MyStyle1"];

            Assert.AreEqual("MyStyle1", style.Name);
            Assert.AreEqual(24, style.Font.Size);
            Assert.AreEqual("Verdana", style.Font.Name);
            Assert.AreEqual(12.0d, style.ParagraphFormat.SpaceAfter);
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
            Document doc = new Document(MyDir + "Style with alias.docx");

            // This document contains a style named "MyStyle,MyStyle Alias 1,MyStyle Alias 2".
            // If a style's name has multiple values separated by commas, each clause is a separate alias.
            Style style = doc.Styles["MyStyle"];
            Assert.AreEqual(new [] { "MyStyle Alias 1", "MyStyle Alias 2" }, style.Aliases);
            Assert.AreEqual("Title", style.BaseStyleName);
            Assert.AreEqual("MyStyle Char", style.LinkedStyleName);

            // We can reference a style using its alias, as well as its name.
            Assert.AreEqual(doc.Styles["MyStyle Alias 1"], doc.Styles["MyStyle Alias 2"]);

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();
            builder.ParagraphFormat.Style = doc.Styles["MyStyle Alias 1"];
            builder.Writeln("Hello world!");
            builder.ParagraphFormat.Style = doc.Styles["MyStyle Alias 2"];
            builder.Write("Hello again!");

            Assert.AreEqual(doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.Style, 
                doc.FirstSection.Body.Paragraphs[1].ParagraphFormat.Style);
            //ExEnd
        }
    }
}