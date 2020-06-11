// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExTxtLoadOptions : ApiExampleBase
    {
        [Test]
        [TestCase(false)]
        [TestCase(true)]
        public void DetectNumberingWithWhitespaces(bool detectNumberingWithWhitespaces)
        {
            //ExStart
            //ExFor:TxtLoadOptions.DetectNumberingWithWhitespaces
            //ExSummary:Shows how lists are detected when plaintext documents are loaded.
            // Create a plaintext document in the form of a string with parts that may be interpreted as lists
            // Upon loading, the first three lists will always be detected by Aspose.Words, and List objects will be created for them after loading
            string textDoc = "Full stop delimiters:\n" +
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

            // The fourth list, with whitespace inbetween the list number and list item contents,
            // will only be detected as a list if "DetectNumberingWithWhitespaces" in a LoadOptions object is set to true,
            // to avoid paragraphs that start with numbers being mistakenly detected as lists
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            loadOptions.DetectNumberingWithWhitespaces = detectNumberingWithWhitespaces;

            // Load the document while applying LoadOptions as a parameter and verify the result
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
        
        [Test]
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

            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                LeadingSpacesOptions = txtLeadingSpacesOptions,
                TrailingSpacesOptions = txtTrailingSpacesOptions
            };

            // Load the document while applying LoadOptions as a parameter and verify the result
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
            //ExSummary:Shows how to detect document direction automatically.
            // Create a LoadOptions object and configure it to detect text direction automatically upon loading
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            loadOptions.DocumentDirection = DocumentDirection.Auto;
 
            // Text like Hebrew/Arabic will be automatically detected as RTL
            Document doc = new Document(MyDir + "Hebrew text.txt", loadOptions);

            Assert.True(doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Bidi);

            doc = new Document(MyDir + "English text.txt", loadOptions);

            Assert.False(doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Bidi);
            //ExEnd
        }
    }
}