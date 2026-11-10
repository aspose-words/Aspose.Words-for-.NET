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
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExCleanupOptions : ApiExampleBase
    {
        [Test]
        public void RemoveUnusedResources()
        {
            //ExStart
            //ExFor:Document.Cleanup(CleanupOptions)
            //ExFor:CleanupOptions
            //ExFor:CleanupOptions.UnusedLists
            //ExFor:CleanupOptions.UnusedStyles
            //ExFor:CleanupOptions.UnusedBuiltinStyles
            //ExSummary:Shows how to remove all unused custom styles from a document. 
            Document doc = new Document();

            doc.Styles.Add(StyleType.List, "MyListStyle1");
            doc.Styles.Add(StyleType.List, "MyListStyle2");
            doc.Styles.Add(StyleType.Character, "MyParagraphStyle1");
            doc.Styles.Add(StyleType.Character, "MyParagraphStyle2");

            // Combined with the built-in styles, the document now has eight styles.
            // A custom style is marked as "used" while there is any text within the document
            // formatted in that style. This means that the 4 styles we added are currently unused.
            Assert.That(doc.Styles.Count, Is.EqualTo(8));

            // Apply a custom character style, and then a custom list style. Doing so will mark them as "used".
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Font.Style = doc.Styles["MyParagraphStyle1"];
            builder.Writeln("Hello world!");

            Aspose.Words.Lists.List docList = doc.Lists.Add(doc.Styles["MyListStyle1"]);
            builder.ListFormat.List = docList;
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");

            // Now, there is one unused character style and one unused list style.
            // The Cleanup() method, when configured with a CleanupOptions object, can target unused styles and remove them.
            CleanupOptions cleanupOptions = new CleanupOptions
            {
                UnusedLists = true, UnusedStyles = true, UnusedBuiltinStyles = true
            };

            doc.Cleanup(cleanupOptions);

            Assert.That(doc.Styles.Count, Is.EqualTo(4));

            // Removing every node that a custom style is applied to marks it as "unused" again. 
            // Rerun the Cleanup method to remove them.
            doc.FirstSection.Body.RemoveAllChildren();
            doc.Cleanup(cleanupOptions);

            Assert.That(doc.Styles.Count, Is.EqualTo(2));
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

            Assert.That(doc.Styles.Count, Is.EqualTo(6));

            // Apply both styles to different paragraphs within the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.ParagraphFormat.StyleName = myStyle.Name;
            builder.Writeln("Hello world!");

            builder.ParagraphFormat.StyleName = duplicateStyle.Name;
            builder.Writeln("Hello again!");

            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            Assert.That(paragraphs[0].ParagraphFormat.Style, Is.EqualTo(myStyle));
            Assert.That(paragraphs[1].ParagraphFormat.Style, Is.EqualTo(duplicateStyle));

            // Configure a CleanOptions object, then call the Cleanup method to substitute all duplicate styles
            // with the original and remove the duplicates from the document.
            CleanupOptions cleanupOptions = new CleanupOptions { DuplicateStyle = true };

            doc.Cleanup(cleanupOptions);

            Assert.That(doc.Styles.Count, Is.EqualTo(5));
            Assert.That(paragraphs[0].ParagraphFormat.Style, Is.EqualTo(myStyle));
            Assert.That(paragraphs[1].ParagraphFormat.Style, Is.EqualTo(myStyle));
            //ExEnd
        }

        [Test]
        public void SymbolicFonts()
        {
            //ExStart:SymbolicFonts
            //GistId:39543adaf4446adfab3bf796a5e45276
            //ExFor:CleanupOptions.SymbolicFonts
            //ExSummary:Shows how to replace text formatted with symbolic fonts (Symbol, Wingdings, etc.) with Unicode text.
            Document doc = new Document(MyDir + "Rendering.docx");

            // Symbolic fonts display characters with glyphs of their own. For example, the "Symbol" font
            // displays the U+F057 character as "Ω", so the text loses its meaning without the font.
            string[] symbolicFonts = { "Symbol", "Wingdings" };
            string Describe(Run run) =>
                $"\"{run.Text}\" ({string.Join(" ", run.Text.Select(c => $"U+{(int)c:X4}"))}), font \"{run.Font.Name}\"";

            List<Paragraph> paragraphs = new List<Paragraph>();
            foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
            {
                if (!symbolicFonts.Contains(run.Font.Name))
                    continue;

                Console.WriteLine($"Before: {Describe(run)}");
                if (!paragraphs.Contains(run.ParentParagraph))
                    paragraphs.Add(run.ParentParagraph);
            }

            // Replace the symbolic text with the equivalent Unicode text formatted with appropriate Unicode fonts.
            CleanupOptions cleanupOptions = new CleanupOptions { SymbolicFonts = true };
            doc.Cleanup(cleanupOptions);

            // The same paragraphs now show the symbols as regular Unicode characters, for example, "Ω" is U+03A9.
            foreach (Paragraph paragraph in paragraphs)
                foreach (Run run in paragraph.Runs)
                    if (run.Text.Any(c => c > 0x7F))
                        Console.WriteLine($"After: {Describe(run)}");

            doc.Save(ArtifactsDir + "CleanupOptions.SymbolicFonts.docx");
            //ExEnd:SymbolicFonts

            Assert.That(paragraphs, Is.Not.Empty);
            Assert.That(doc.GetChildNodes(NodeType.Run, true).Cast<Run>().Any(r => symbolicFonts.Contains(r.Font.Name)), Is.False);
        }
    }
}
