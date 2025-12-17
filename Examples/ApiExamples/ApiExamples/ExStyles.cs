// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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

            Assert.That(doc.Styles.Count, Is.EqualTo(4));

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

                    Assert.That(curStyle.Document, Is.EqualTo(doc));
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
            //ExFor:Style.AutomaticallyUpdate
            //ExSummary:Shows how to create and apply a custom style.
            Document doc = new Document();

            Style style = doc.Styles.Add(StyleType.Paragraph, "MyStyle");
            style.Font.Name = "Times New Roman";
            style.Font.Size = 16;
            style.Font.Color = Color.Navy;
            // Automatically redefine style.
            style.AutomaticallyUpdate = true;

            DocumentBuilder builder = new DocumentBuilder(doc);

            // Apply one of the styles from the document to the paragraph that the document builder is creating.
            builder.ParagraphFormat.Style = doc.Styles["MyStyle"];
            builder.Writeln("Hello world!");

            Style firstParagraphStyle = doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Style;

            Assert.That(firstParagraphStyle, Is.EqualTo(style));

            // Remove our custom style from the document's styles collection.
            doc.Styles["MyStyle"].Remove();

            firstParagraphStyle = doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Style;

            // Any text that used a removed style reverts to the default formatting.
            Assert.That(doc.Styles.Any(s => s.Name == "MyStyle"), Is.False);
            Assert.That(firstParagraphStyle.Font.Name, Is.EqualTo("Times New Roman"));
            Assert.That(firstParagraphStyle.Font.Size, Is.EqualTo(12.0d));
            Assert.That(firstParagraphStyle.Font.Color.ToArgb(), Is.EqualTo(Color.Empty.ToArgb()));
            //ExEnd
        }

        [Test]
        public void StyleCollection()
        {
            //ExStart
            //ExFor:StyleCollection.Add(StyleType,String)
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

            Assert.That(styles[4].Font.Name, Is.EqualTo("Courier New"));
            Assert.That(styles["MyStyle"].ParagraphFormat.FirstLineIndent, Is.EqualTo(15.0));
            //ExEnd
        }

        [Test]
        public void RemoveStylesFromStyleGallery()
        {
            //ExStart
            //ExFor:StyleCollection.ClearQuickStyleGallery
            //ExSummary:Shows how to remove styles from Style Gallery panel.
            Document doc = new Document();
            // Note that remove styles work only with DOCX format for now.
            doc.Styles.ClearQuickStyleGallery();

            doc.Save(ArtifactsDir + "Styles.RemoveStylesFromStyleGallery.docx");
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
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
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

            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
                if (para.ParagraphFormat.Style.StyleIdentifier >= StyleIdentifier.Toc1 &&
                    para.ParagraphFormat.Style.StyleIdentifier <= StyleIdentifier.Toc9)
                {
                    TabStop tabStop = para.GetEffectiveTabStops()[0];
                    Assert.That(tabStop.Position, Is.EqualTo(400.8d));
                    Assert.That(tabStop.Alignment, Is.EqualTo(TabAlignment.Right));
                    Assert.That(tabStop.Leader, Is.EqualTo(TabLeader.Dots));
                }
        }

        [Test]
        public void CopyStyleSameDocument()
        {
            //ExStart
            //ExFor:StyleCollection.AddCopy(Style)
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

            Assert.That(newStyle.Name, Is.EqualTo("My Heading 1"));
            Assert.That(doc.Styles["Heading 1"].Name, Is.EqualTo("Heading 1"));

            Assert.That(newStyle.Type, Is.EqualTo(doc.Styles["Heading 1"].Type));
            Assert.That(newStyle.Font.Name, Is.EqualTo(doc.Styles["Heading 1"].Font.Name));
            Assert.That(newStyle.Font.Size, Is.EqualTo(doc.Styles["Heading 1"].Font.Size));
            Assert.That(newStyle.Font.Color, Is.Not.EqualTo(doc.Styles["Heading 1"].Font.Color));
            //ExEnd
        }

        [Test]
        public void CopyStyleDifferentDocument()
        {
            //ExStart
            //ExFor:StyleCollection.AddCopy(Style)
            //ExSummary:Shows how to import a style from one document into a different document.
            Document srcDoc = new Document();

            // Create a custom style for the source document.
            Style srcStyle = srcDoc.Styles.Add(StyleType.Paragraph, "MyStyle");
            srcStyle.Font.Color = Color.Red;

            // Import the source document's custom style into the destination document.
            Document dstDoc = new Document();
            Style newStyle = dstDoc.Styles.AddCopy(srcStyle);

            // The imported style has an appearance identical to its source style.
            Assert.That(newStyle.Name, Is.EqualTo("MyStyle"));
            Assert.That(newStyle.Font.Color.ToArgb(), Is.EqualTo(Color.Red.ToArgb()));
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

            Assert.That(doc.Styles.DefaultFont.Bold, Is.True);
            Assert.That(doc.Styles.DefaultFont.Name, Is.EqualTo("PMingLiU"));
            Assert.That(doc.Styles.DefaultParagraphFormat.SpaceAfter, Is.EqualTo(20));
            Assert.That(doc.Styles.DefaultParagraphFormat.Alignment, Is.EqualTo(ParagraphAlignment.Right));
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

            Assert.That(style.Name, Is.EqualTo("MyStyle1"));
            Assert.That(style.Font.Size, Is.EqualTo(24));
            Assert.That(style.Font.Name, Is.EqualTo("Verdana"));
            Assert.That(style.ParagraphFormat.SpaceAfter, Is.EqualTo(12.0d));
        }

        [Test]
        public void StyleAliases()
        {
            //ExStart
            //ExFor:Style.Aliases
            //ExFor:Style.BaseStyleName
            //ExFor:Style.Equals(Style)
            //ExFor:Style.LinkedStyleName
            //ExSummary:Shows how to use style aliases.
            Document doc = new Document(MyDir + "Style with alias.docx");

            // This document contains a style named "MyStyle,MyStyle Alias 1,MyStyle Alias 2".
            // If a style's name has multiple values separated by commas, each clause is a separate alias.
            Style style = doc.Styles["MyStyle"];
            Assert.That(style.Aliases, Is.EqualTo(new [] { "MyStyle Alias 1", "MyStyle Alias 2" }));
            Assert.That(style.BaseStyleName, Is.EqualTo("Title"));
            Assert.That(style.LinkedStyleName, Is.EqualTo("MyStyle Char"));

            // We can reference a style using its alias, as well as its name.
            Assert.That(doc.Styles["MyStyle Alias 2"], Is.EqualTo(doc.Styles["MyStyle Alias 1"]));

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();
            builder.ParagraphFormat.Style = doc.Styles["MyStyle Alias 1"];
            builder.Writeln("Hello world!");
            builder.ParagraphFormat.Style = doc.Styles["MyStyle Alias 2"];
            builder.Write("Hello again!");

            Assert.That(doc.FirstSection.Body.Paragraphs[1].ParagraphFormat.Style, Is.EqualTo(doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.Style));
            //ExEnd
        }

        [Test]
        public void LatentStyles()
        {
            // This test is to check that after re-saving a document it doesn't lose LatentStyle information
            // for 4 styles from documents created in Microsoft Word.
            Document doc = new Document(MyDir + "Blank.docx");

            doc.Save(ArtifactsDir + "Styles.LatentStyles.docx");

            TestUtil.DocPackageFileContainsString(
                "<w:lsdException w:name=\"Mention\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\" />",
                ArtifactsDir + "Styles.LatentStyles.docx", "styles.xml");
            TestUtil.DocPackageFileContainsString(
                "<w:lsdException w:name=\"Smart Hyperlink\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\" />",
                ArtifactsDir + "Styles.LatentStyles.docx", "styles.xml");
            TestUtil.DocPackageFileContainsString(
                "<w:lsdException w:name=\"Hashtag\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\" />",
                ArtifactsDir + "Styles.LatentStyles.docx", "styles.xml");
            TestUtil.DocPackageFileContainsString(
                "<w:lsdException w:name=\"Unresolved Mention\" w:semiHidden=\"1\" w:unhideWhenUsed=\"1\" />",
                ArtifactsDir + "Styles.LatentStyles.docx", "styles.xml");
        }

        [Test]
        public void LockStyle()
        {
            //ExStart:LockStyle
            //GistId:3428e84add5beb0d46a8face6e5fc858
            //ExFor:Style.Locked
            //ExSummary:Shows how to lock style.
            Document doc = new Document();

            Style styleHeading1 = doc.Styles[StyleIdentifier.Heading1];
            if (!styleHeading1.Locked)
                styleHeading1.Locked = true;

            doc.Save(ArtifactsDir + "Styles.LockStyle.docx");
            //ExEnd:LockStyle

            doc = new Document(ArtifactsDir + "Styles.LockStyle.docx");
            Assert.That(doc.Styles[StyleIdentifier.Heading1].Locked, Is.True);
        }

        [Test]
        public void StylePriority()
        {
            //ExStart:StylePriority
            //GistId:a775441ecb396eea917a2717cb9e8f8f
            //ExFor:Style.Priority
            //ExFor:Style.UnhideWhenUsed
            //ExFor:Style.SemiHidden
            //ExSummary:Shows how to prioritize and hide a style.
            Document doc = new Document();
            Style styleTitle = doc.Styles[StyleIdentifier.Subtitle];
            
            if (styleTitle.Priority == 9)
                styleTitle.Priority = 10;

            if (!styleTitle.UnhideWhenUsed)
                styleTitle.UnhideWhenUsed = true;

            if (styleTitle.SemiHidden)
                styleTitle.SemiHidden = true;

            doc.Save(ArtifactsDir + "Styles.StylePriority.docx");
            //ExEnd:StylePriority
        }

        [Test]
        public void LinkedStyleName()
        {
            //ExStart:LinkedStyleName
            //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
            //ExFor:Style.LinkedStyleName
            //ExSummary:Shows how to link styles among themselves.
            Document doc = new Document();

            Style styleHeading1 = doc.Styles[StyleIdentifier.Heading1];

            Style styleHeading1Char = doc.Styles.Add(StyleType.Character, "Heading 1 Char");
            styleHeading1Char.Font.Name = "Verdana";
            styleHeading1Char.Font.Bold = true;
            styleHeading1Char.Font.Border.LineStyle = LineStyle.Dot;
            styleHeading1Char.Font.Border.LineWidth = 15;

            styleHeading1.LinkedStyleName = "Heading 1 Char";

            Assert.That(styleHeading1.LinkedStyleName, Is.EqualTo("Heading 1 Char"));
            Assert.That(styleHeading1Char.LinkedStyleName, Is.EqualTo("Heading 1"));
            //ExEnd:LinkedStyleName
        }
    }
}
