// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
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
                Assert.AreEqual(4, doc.Lists.Count);
                Assert.True(doc.FirstSection.Body.Paragraphs.Any(p => p.GetText().Contains("Fourth list") && ((Paragraph)p).IsListItem));
            }
            else
            {
                Assert.AreEqual(3, doc.Lists.Count);
                Assert.False(doc.FirstSection.Body.Paragraphs.Any(p => p.GetText().Contains("Fourth list") && ((Paragraph)p).IsListItem));
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
                    Assert.AreEqual(37.8d, paragraphs[0].ParagraphFormat.FirstLineIndent);
                    Assert.AreEqual(25.2d, paragraphs[1].ParagraphFormat.FirstLineIndent);
                    Assert.AreEqual(6.3d, paragraphs[2].ParagraphFormat.FirstLineIndent);

                    Assert.True(paragraphs[0].GetText().StartsWith("Line 1"));
                    Assert.True(paragraphs[1].GetText().StartsWith("Line 2"));
                    Assert.True(paragraphs[2].GetText().StartsWith("Line 3"));
                    break;
                case TxtLeadingSpacesOptions.Preserve:
                    Assert.True(paragraphs.All(p => ((Paragraph)p).ParagraphFormat.FirstLineIndent == 0.0d));

                    Assert.True(paragraphs[0].GetText().StartsWith("      Line 1"));
                    Assert.True(paragraphs[1].GetText().StartsWith("    Line 2"));
                    Assert.True(paragraphs[2].GetText().StartsWith(" Line 3"));
                    break;
                case TxtLeadingSpacesOptions.Trim:
                    Assert.True(paragraphs.All(p => ((Paragraph)p).ParagraphFormat.FirstLineIndent == 0.0d));

                    Assert.True(paragraphs[0].GetText().StartsWith("Line 1"));
                    Assert.True(paragraphs[1].GetText().StartsWith("Line 2"));
                    Assert.True(paragraphs[2].GetText().StartsWith("Line 3"));
                    break;
            }

            switch (txtTrailingSpacesOptions)
            {
                case TxtTrailingSpacesOptions.Preserve:
                    Assert.True(paragraphs[0].GetText().EndsWith("Line 1 \r"));
                    Assert.True(paragraphs[1].GetText().EndsWith("Line 2   \r"));
                    Assert.True(paragraphs[2].GetText().EndsWith("Line 3       \f"));
                    break;
                case TxtTrailingSpacesOptions.Trim:
                    Assert.True(paragraphs[0].GetText().EndsWith("Line 1\r"));
                    Assert.True(paragraphs[1].GetText().EndsWith("Line 2\r"));
                    Assert.True(paragraphs[2].GetText().EndsWith("Line 3\f"));
                    break;
            }
            //ExEnd
        }

        [Test]
        public void DetectDocumentDirection()
        {
            //ExStart
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

            Assert.True(doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Bidi);

            // Detect English text as right-to-left.
            doc = new Document(MyDir + "English text.txt", loadOptions);

            Assert.False(doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Bidi);
            //ExEnd
        }
    }
}