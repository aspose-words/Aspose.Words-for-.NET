﻿// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Loading;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExTxtLoadOptions : ApiExampleBase
    {
        [TestCase(false)]
        [TestCase(true)]
        public void DetectNumberingWithWhitespaces(bool detectNumberingWithWhitespaces)
        {
            //ExStart
            //ExFor:TxtLoadOptions.DetectNumberingWithWhitespaces
            //ExSummary:Shows how to detect lists when loading plaintext documents.
            // Create a plaintext document in a string with four separate parts that we may interpret as lists,
            // with different delimiters. Upon loading the plaintext document into a "Document" object,
            // Aspose.Words will always detect the first three lists and will add a "List" object
            // for each to the document's "Lists" property.
            const string textDoc = "Full stop delimiters:\n" +
                                   "1. First list item 1\n" +
                                   "2. First list item 2\n" +
                                   "3. First list item 3\n\n" +
                                   "Right bracket delimiters:\n" +
                                   "1) Second list item 1\n" +
                                   "2) Second list item 2\n" +
                                   "3) Second list item 3\n\n" +
                                   "Bullet delimiters:\n" +
                                   "• Third list item 1\n" +
                                   "• Third list item 2\n" +
                                   "• Third list item 3\n\n" +
                                   "Whitespace delimiters:\n" +
                                   "1 Fourth list item 1\n" +
                                   "2 Fourth list item 2\n" +
                                   "3 Fourth list item 3";

            // Create a "TxtLoadOptions" object, which we can pass to a document's constructor
            // to modify how we load a plaintext document.
            TxtLoadOptions loadOptions = new TxtLoadOptions();

            // Set the "DetectNumberingWithWhitespaces" property to "true" to detect numbered items
            // with whitespace delimiters, such as the fourth list in our document, as lists.
            // This may also falsely detect paragraphs that begin with numbers as lists.
            // Set the "DetectNumberingWithWhitespaces" property to "false"
            // to not create lists from numbered items with whitespace delimiters.
            loadOptions.DetectNumberingWithWhitespaces = detectNumberingWithWhitespaces;

            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(textDoc)), loadOptions);

            if (detectNumberingWithWhitespaces)
            {
                Assert.That(doc.Lists.Count, Is.EqualTo(4));
                Assert.That(doc.FirstSection.Body.Paragraphs.Any(p => p.GetText().Contains("Fourth list") && ((Paragraph)p).IsListItem), Is.True);
            }
            else
            {
                Assert.That(doc.Lists.Count, Is.EqualTo(3));
                Assert.That(doc.FirstSection.Body.Paragraphs.Any(p => p.GetText().Contains("Fourth list") && ((Paragraph)p).IsListItem), Is.False);
            }
            //ExEnd
        }
        
        [TestCase(TxtLeadingSpacesOptions.Preserve, TxtTrailingSpacesOptions.Preserve)]
        [TestCase(TxtLeadingSpacesOptions.ConvertToIndent, TxtTrailingSpacesOptions.Preserve)]
        [TestCase(TxtLeadingSpacesOptions.Trim, TxtTrailingSpacesOptions.Trim)]
        public void TrailSpaces(TxtLeadingSpacesOptions txtLeadingSpacesOptions, TxtTrailingSpacesOptions txtTrailingSpacesOptions)
        {
            //ExStart
            //ExFor:TxtLoadOptions.TrailingSpacesOptions
            //ExFor:TxtLoadOptions.LeadingSpacesOptions
            //ExFor:TxtTrailingSpacesOptions
            //ExFor:TxtLeadingSpacesOptions
            //ExSummary:Shows how to trim whitespace when loading plaintext documents.
            string textDoc = "      Line 1 \n" +
                             "    Line 2   \n" +
                             " Line 3       ";

            // Create a "TxtLoadOptions" object, which we can pass to a document's constructor
            // to modify how we load a plaintext document.
            TxtLoadOptions loadOptions = new TxtLoadOptions();

            // Set the "LeadingSpacesOptions" property to "TxtLeadingSpacesOptions.Preserve"
            // to preserve all whitespace characters at the start of every line.
            // Set the "LeadingSpacesOptions" property to "TxtLeadingSpacesOptions.ConvertToIndent"
            // to remove all whitespace characters from the start of every line,
            // and then apply a left first line indent to the paragraph to simulate the effect of the whitespaces.
            // Set the "LeadingSpacesOptions" property to "TxtLeadingSpacesOptions.Trim"
            // to remove all whitespace characters from every line's start.
            loadOptions.LeadingSpacesOptions = txtLeadingSpacesOptions;

            // Set the "TrailingSpacesOptions" property to "TxtTrailingSpacesOptions.Preserve"
            // to preserve all whitespace characters at the end of every line. 
            // Set the "TrailingSpacesOptions" property to "TxtTrailingSpacesOptions.Trim" to 
            // remove all whitespace characters from the end of every line.
            loadOptions.TrailingSpacesOptions = txtTrailingSpacesOptions;

            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(textDoc)), loadOptions);
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            switch (txtLeadingSpacesOptions)
            {
                case TxtLeadingSpacesOptions.ConvertToIndent:
                    Assert.That(paragraphs[0].ParagraphFormat.FirstLineIndent, Is.EqualTo(37.8d));
                    Assert.That(paragraphs[1].ParagraphFormat.FirstLineIndent, Is.EqualTo(25.2d));
                    Assert.That(paragraphs[2].ParagraphFormat.FirstLineIndent, Is.EqualTo(6.3d));

                    Assert.That(paragraphs[0].GetText().StartsWith("Line 1"), Is.True);
                    Assert.That(paragraphs[1].GetText().StartsWith("Line 2"), Is.True);
                    Assert.That(paragraphs[2].GetText().StartsWith("Line 3"), Is.True);
                    break;
                case TxtLeadingSpacesOptions.Preserve:
                    Assert.That(paragraphs.All(p => ((Paragraph)p).ParagraphFormat.FirstLineIndent == 0.0d), Is.True);

                    Assert.That(paragraphs[0].GetText().StartsWith("      Line 1"), Is.True);
                    Assert.That(paragraphs[1].GetText().StartsWith("    Line 2"), Is.True);
                    Assert.That(paragraphs[2].GetText().StartsWith(" Line 3"), Is.True);
                    break;
                case TxtLeadingSpacesOptions.Trim:
                    Assert.That(paragraphs.All(p => ((Paragraph)p).ParagraphFormat.FirstLineIndent == 0.0d), Is.True);

                    Assert.That(paragraphs[0].GetText().StartsWith("Line 1"), Is.True);
                    Assert.That(paragraphs[1].GetText().StartsWith("Line 2"), Is.True);
                    Assert.That(paragraphs[2].GetText().StartsWith("Line 3"), Is.True);
                    break;
            }

            switch (txtTrailingSpacesOptions)
            {
                case TxtTrailingSpacesOptions.Preserve:
                    Assert.That(paragraphs[0].GetText().EndsWith("Line 1 \r"), Is.True);
                    Assert.That(paragraphs[1].GetText().EndsWith("Line 2   \r"), Is.True);
                    Assert.That(paragraphs[2].GetText().EndsWith("Line 3       \f"), Is.True);
                    break;
                case TxtTrailingSpacesOptions.Trim:
                    Assert.That(paragraphs[0].GetText().EndsWith("Line 1\r"), Is.True);
                    Assert.That(paragraphs[1].GetText().EndsWith("Line 2\r"), Is.True);
                    Assert.That(paragraphs[2].GetText().EndsWith("Line 3\f"), Is.True);
                    break;
            }
            //ExEnd
        }

        [Test]
        public void DetectDocumentDirection()
        {
            //ExStart
            //ExFor:DocumentDirection
            //ExFor:TxtLoadOptions.DocumentDirection
            //ExFor:ParagraphFormat.Bidi
            //ExSummary:Shows how to detect plaintext document text direction.
            // Create a "TxtLoadOptions" object, which we can pass to a document's constructor
            // to modify how we load a plaintext document.
            TxtLoadOptions loadOptions = new TxtLoadOptions();

            // Set the "DocumentDirection" property to "DocumentDirection.Auto" automatically detects
            // the direction of every paragraph of text that Aspose.Words loads from plaintext.
            // Each paragraph's "Bidi" property will store its direction.
            loadOptions.DocumentDirection = DocumentDirection.Auto;
 
            // Detect Hebrew text as right-to-left.
            Document doc = new Document(MyDir + "Hebrew text.txt", loadOptions);

            Assert.That(doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Bidi, Is.True);

            // Detect English text as right-to-left.
            doc = new Document(MyDir + "English text.txt", loadOptions);

            Assert.That(doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Bidi, Is.False);
            //ExEnd
        }

        [Test]
        public void AutoNumberingDetection()
        {
            //ExStart
            //ExFor:TxtLoadOptions.AutoNumberingDetection
            //ExSummary:Shows how to disable automatic numbering detection.
            TxtLoadOptions options = new TxtLoadOptions { AutoNumberingDetection = false };
            Document doc = new Document(MyDir + "Number detection.txt", options);
            //ExEnd

            int listItemsCount = 0;
            foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (paragraph.IsListItem)
                    listItemsCount++;
            }

            Assert.That(listItemsCount, Is.EqualTo(0));
        }

        [Test]
        public void DetectHyperlinks()
        {
            //ExStart:DetectHyperlinks
            //GistId:3428e84add5beb0d46a8face6e5fc858
            //ExFor:TxtLoadOptions
            //ExFor:TxtLoadOptions.#ctor
            //ExFor:TxtLoadOptions.DetectHyperlinks
            //ExSummary:Shows how to read and display hyperlinks.
            const string inputText = "Some links in TXT:\n" +
                    "https://www.aspose.com/\n" +
                    "https://docs.aspose.com/words/net/\n";

            using (Stream stream = new MemoryStream())
            {
                byte[] buf = Encoding.ASCII.GetBytes(inputText);
                stream.Write(buf, 0, buf.Length);

                // Load document with hyperlinks.
                Document doc = new Document(stream, new TxtLoadOptions() { DetectHyperlinks = true });

                // Print hyperlinks text.
                foreach (Field field in doc.Range.Fields)
                    Console.WriteLine(field.Result);

                Assert.That("https://www.aspose.com/", Is.EqualTo(doc.Range.Fields[0].Result.Trim()));
                Assert.That("https://docs.aspose.com/words/net/", Is.EqualTo(doc.Range.Fields[1].Result.Trim()));
            }
            //ExEnd:DetectHyperlinks
        }
    }
}
