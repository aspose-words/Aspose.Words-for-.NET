// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
    class ExCleanupOptions : ApiExampleBase
    {
        [Test]
        public void RemoveUnusedResources()
        {
            //ExStart
            //ExFor:Document.Cleanup(CleanupOptions)
            //ExFor:CleanupOptions
            //ExFor:CleanupOptions.UnusedLists
            //ExFor:CleanupOptions.UnusedStyles
            //ExSummary:Shows how to remove all unused custom styles from a document. 
            Document doc = new Document();

            doc.Styles.Add(StyleType.List, "MyListStyle1");
            doc.Styles.Add(StyleType.List, "MyListStyle2");
            doc.Styles.Add(StyleType.Character, "MyParagraphStyle1");
            doc.Styles.Add(StyleType.Character, "MyParagraphStyle2");

            // Combined with the built-in styles, the document now has eight styles.
            // A custom style is marked as "used" while there is any text within the document
            // formatted in that style. This means that the 4 styles we added are currently unused.
            Assert.AreEqual(8, doc.Styles.Count);

            // Apply a custom character style, and then a custom list style. Doing so will mark them as "used".
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Font.Style = doc.Styles["MyParagraphStyle1"];
            builder.Writeln("Hello world!");

            Aspose.Words.Lists.List list = doc.Lists.Add(doc.Styles["MyListStyle1"]);
            builder.ListFormat.List = list;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");

            // Now, there is one unused character style and one unused list style.
            // The Cleanup() method, when configured with a CleanupOptions object, can target unused styles and remove them.
            CleanupOptions cleanupOptions = new CleanupOptions();
            cleanupOptions.UnusedLists = true;
            cleanupOptions.UnusedStyles = true;

            doc.Cleanup(cleanupOptions);

            Assert.AreEqual(6, doc.Styles.Count);

            // Removing every node that a custom style is applied to marks it as "unused" again. 
            // Rerun the Cleanup method to remove them.
            doc.FirstSection.Body.RemoveAllChildren();
            doc.Cleanup(cleanupOptions);

            Assert.AreEqual(4, doc.Styles.Count);
            //ExEnd
        }

        [Test]
        public void RemoveDuplicateStyles()
        {
            //ExStart
            //ExFor:CleanupOptions.DuplicateStyle
            //ExSummary:Shows how to remove duplicated styles from the document.
            Document doc = new Document();

            // Add two styles to the document with identical properties,
            // but different names. The second style is considered a duplicate of the first.
            Style myStyle = doc.Styles.Add(StyleType.Paragraph, "MyStyle1");
            myStyle.Font.Size = 14;
            myStyle.Font.Name = "Courier New";
            myStyle.Font.Color = Color.Blue;

            Style duplicateStyle = doc.Styles.Add(StyleType.Paragraph, "MyStyle2");
            duplicateStyle.Font.Size = 14;
            duplicateStyle.Font.Name = "Courier New";
            duplicateStyle.Font.Color = Color.Blue;

            Assert.AreEqual(6, doc.Styles.Count);

            // Apply both styles to different paragraphs within the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.ParagraphFormat.StyleName = myStyle.Name;
            builder.Writeln("Hello world!");

            builder.ParagraphFormat.StyleName = duplicateStyle.Name;
            builder.Writeln("Hello again!");

            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            Assert.AreEqual(myStyle, paragraphs[0].ParagraphFormat.Style);
            Assert.AreEqual(duplicateStyle, paragraphs[1].ParagraphFormat.Style);

            // Configure a CleanOptions object, then call the Cleanup method to substitute all duplicate styles
            // with the original and remove the duplicates from the document.
            CleanupOptions cleanupOptions = new CleanupOptions();
            cleanupOptions.DuplicateStyle = true;

            doc.Cleanup(cleanupOptions);

            Assert.AreEqual(5, doc.Styles.Count);
            Assert.AreEqual(myStyle, paragraphs[0].ParagraphFormat.Style);
            Assert.AreEqual(myStyle, paragraphs[1].ParagraphFormat.Style);
            //ExEnd
        }
    }
}
