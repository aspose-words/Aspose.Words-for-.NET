// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExTxtSaveOptions : ApiExampleBase
    {
        [TestCase(false)]
        [TestCase(true)]
        public void PageBreaks(bool forcePageBreaks)
        {
            //ExStart
            //ExFor:TxtSaveOptionsBase.ForcePageBreaks
            //ExSummary:Shows how to specify whether to preserve page breaks when exporting a document to plaintext.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 3");

            // Create a "TxtSaveOptions" object, which we can pass to the document's "Save"
            // method to modify how we save the document to plaintext.
            TxtSaveOptions saveOptions = new TxtSaveOptions();

            // The Aspose.Words "Document" objects have page breaks, just like Microsoft Word documents.
            // Save formats such as ".txt" are one continuous body of text without page breaks.
            // Set the "ForcePageBreaks" property to "true" to preserve all page breaks in the form of '\f' characters.
            // Set the "ForcePageBreaks" property to "false" to discard all page breaks.
            saveOptions.ForcePageBreaks = forcePageBreaks;

            doc.Save(ArtifactsDir + "TxtSaveOptions.PageBreaks.txt", saveOptions);
            
            // If we load a plaintext document with page breaks,
            // the "Document" object will use them to split the body into pages.
            doc = new Document(ArtifactsDir + "TxtSaveOptions.PageBreaks.txt");

            Assert.AreEqual(forcePageBreaks ? 3 : 1, doc.PageCount);
            //ExEnd

            TestUtil.FileContainsString(
                forcePageBreaks ? "Page 1\r\n\fPage 2\r\n\fPage 3\r\n\r\n" : "Page 1\r\nPage 2\r\nPage 3\r\n\r\n",
                ArtifactsDir + "TxtSaveOptions.PageBreaks.txt");
        }

        [TestCase(false)]
        [TestCase(true)]
        public void AddBidiMarks(bool addBidiMarks)
        {
            //ExStart
            //ExFor:TxtSaveOptions.AddBidiMarks
            //ExSummary:Shows how to insert Unicode Character 'RIGHT-TO-LEFT MARK' (U+200F) before each bi-directional Run in text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");
            builder.ParagraphFormat.Bidi = true;
            builder.Writeln("שלום עולם!");
            builder.Writeln("مرحبا بالعالم!");

            // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
            // to modify how we save the document to plaintext.
            TxtSaveOptions saveOptions = new TxtSaveOptions { Encoding = System.Text.Encoding.Unicode};

            // Set the "AddBidiMarks" property to "true" to add marks before runs
            // with right-to-left text to indicate the fact.
            // Set the "AddBidiMarks" property to "false" to write all left-to-right
            // and right-to-left run equally with nothing to indicate which is which.
            saveOptions.AddBidiMarks = addBidiMarks;

            doc.Save(ArtifactsDir + "TxtSaveOptions.AddBidiMarks.txt", saveOptions);

            string docText = System.Text.Encoding.Unicode.GetString(File.ReadAllBytes(ArtifactsDir + "TxtSaveOptions.AddBidiMarks.txt"));

            if (addBidiMarks)
            {
                Assert.AreEqual("\uFEFFHello world!‎\r\nשלום עולם!‏\r\nمرحبا بالعالم!‏\r\n\r\n", docText);
                Assert.True(docText.Contains("\u200f"));
            }
            else
            {
                Assert.AreEqual("\uFEFFHello world!\r\nשלום עולם!\r\nمرحبا بالعالم!\r\n\r\n", docText);
                Assert.False(docText.Contains("\u200f"));
            }
            //ExEnd
        }

        [TestCase(TxtExportHeadersFootersMode.AllAtEnd)]
        [TestCase(TxtExportHeadersFootersMode.PrimaryOnly)]
        [TestCase(TxtExportHeadersFootersMode.None)]
        public void ExportHeadersFooters(TxtExportHeadersFootersMode txtExportHeadersFootersMode)
        {
            //ExStart
            //ExFor:TxtSaveOptionsBase.ExportHeadersFootersMode
            //ExFor:TxtExportHeadersFootersMode
            //ExSummary:Shows how to specify how to export headers and footers to plain text format.
            Document doc = new Document();

            // Insert even and primary headers/footers into the document.
            // The primary header/footers will override the even headers/footers.
            doc.FirstSection.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.HeaderEven));
            doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderEven].AppendParagraph("Even header");
            doc.FirstSection.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.FooterEven));
            doc.FirstSection.HeadersFooters[HeaderFooterType.FooterEven].AppendParagraph("Even footer");
            doc.FirstSection.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.HeaderPrimary));
            doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].AppendParagraph("Primary header");
            doc.FirstSection.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.FooterPrimary));
            doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary].AppendParagraph("Primary footer");

            // Insert pages to display these headers and footers.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2");
            builder.InsertBreak(BreakType.PageBreak); 
            builder.Write("Page 3");

            // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
            // to modify how we save the document to plaintext.
            TxtSaveOptions saveOptions = new TxtSaveOptions();

            // Set the "ExportHeadersFootersMode" property to "TxtExportHeadersFootersMode.None"
            // to not export any headers/footers.
            // Set the "ExportHeadersFootersMode" property to "TxtExportHeadersFootersMode.PrimaryOnly"
            // to only export primary headers/footers.
            // Set the "ExportHeadersFootersMode" property to "TxtExportHeadersFootersMode.AllAtEnd"
            // to place all headers and footers for all section bodies at the end of the document.
            saveOptions.ExportHeadersFootersMode = txtExportHeadersFootersMode;

            doc.Save(ArtifactsDir + "TxtSaveOptions.ExportHeadersFooters.txt", saveOptions);

            string docText = File.ReadAllText(ArtifactsDir + "TxtSaveOptions.ExportHeadersFooters.txt");

            switch (txtExportHeadersFootersMode)
            {
                case TxtExportHeadersFootersMode.AllAtEnd:
                    Assert.AreEqual("Page 1\r\n" +
                                    "Page 2\r\n" +
                                    "Page 3\r\n" +
                                    "Even header\r\n\r\n" +
                                    "Primary header\r\n\r\n" +
                                    "Even footer\r\n\r\n" +
                                    "Primary footer\r\n\r\n", docText);
                    break;
                case TxtExportHeadersFootersMode.PrimaryOnly:
                    Assert.AreEqual("Primary header\r\n" +
                                    "Page 1\r\n" +
                                    "Page 2\r\n" +
                                    "Page 3\r\n" +
                                    "Primary footer\r\n", docText);
                    break;
                case TxtExportHeadersFootersMode.None:
                    Assert.AreEqual("Page 1\r\n" +
                                    "Page 2\r\n" +
                                    "Page 3\r\n", docText);
                    break;
            }
            //ExEnd
        }

        [Test]
        public void TxtListIndentation()
        {
            //ExStart
            //ExFor:TxtListIndentation
            //ExFor:TxtListIndentation.Count
            //ExFor:TxtListIndentation.Character
            //ExFor:TxtSaveOptions.ListIndentation
            //ExSummary:Shows how to configure list indenting when saving a document to plaintext.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a list with three levels of indentation.
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Item 1");
            builder.ListFormat.ListIndent();
            builder.Writeln("Item 2");
            builder.ListFormat.ListIndent(); 
            builder.Write("Item 3");

            // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
            // to modify how we save the document to plaintext.
            TxtSaveOptions txtSaveOptions = new TxtSaveOptions();

            // Set the "Character" property to assign a character to use
            // for padding that simulates list indentation in plaintext.
            txtSaveOptions.ListIndentation.Character = ' ';

            // Set the "Count" property to specify the number of times
            // to place the padding character for each list indent level.
            txtSaveOptions.ListIndentation.Count = 3;

            doc.Save(ArtifactsDir + "TxtSaveOptions.TxtListIndentation.txt", txtSaveOptions);

            string docText = File.ReadAllText(ArtifactsDir + "TxtSaveOptions.TxtListIndentation.txt");

            Assert.AreEqual("1. Item 1\r\n" +
                            "   a. Item 2\r\n" +
                            "      i. Item 3\r\n", docText);
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void SimplifyListLabels(bool simplifyListLabels)
        {
            //ExStart
            //ExFor:TxtSaveOptions.SimplifyListLabels
            //ExSummary:Shows how to change the appearance of lists when saving a document to plaintext.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a bulleted list with five levels of indentation.
            builder.ListFormat.ApplyBulletDefault();
            builder.Writeln("Item 1");
            builder.ListFormat.ListIndent();
            builder.Writeln("Item 2");
            builder.ListFormat.ListIndent();
            builder.Writeln("Item 3");
            builder.ListFormat.ListIndent();
            builder.Writeln("Item 4");
            builder.ListFormat.ListIndent();
            builder.Write("Item 5");

            // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
            // to modify how we save the document to plaintext.
            TxtSaveOptions txtSaveOptions = new TxtSaveOptions();

            // Set the "SimplifyListLabels" property to "true" to convert some list
            // symbols into simpler ASCII characters, such as '*', 'o', '+', '>', etc.
            // Set the "SimplifyListLabels" property to "false" to preserve as many original list symbols as possible.
            txtSaveOptions.SimplifyListLabels = simplifyListLabels;

            doc.Save(ArtifactsDir + "TxtSaveOptions.SimplifyListLabels.txt", txtSaveOptions);

            string docText = File.ReadAllText(ArtifactsDir + "TxtSaveOptions.SimplifyListLabels.txt");

            if (simplifyListLabels)
                Assert.AreEqual("* Item 1\r\n" +
                                "  > Item 2\r\n" +
                                "    + Item 3\r\n" +
                                "      - Item 4\r\n" +
                                "        o Item 5\r\n", docText);
            else
                Assert.AreEqual("· Item 1\r\n" +
                                "o Item 2\r\n" +
                                "§ Item 3\r\n" +
                                "· Item 4\r\n" +
                                "o Item 5\r\n", docText);
            //ExEnd
        }

        [Test]
        public void ParagraphBreak()
        {
            //ExStart
            //ExFor:TxtSaveOptions
            //ExFor:TxtSaveOptions.SaveFormat
            //ExFor:TxtSaveOptionsBase
            //ExFor:TxtSaveOptionsBase.ParagraphBreak
            //ExSummary:Shows how to save a .txt document with a custom paragraph break.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Paragraph 1.");
            builder.Writeln("Paragraph 2.");
            builder.Write("Paragraph 3.");

            // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
            // to modify how we save the document to plaintext.
            TxtSaveOptions txtSaveOptions = new TxtSaveOptions();

            Assert.AreEqual(SaveFormat.Text, txtSaveOptions.SaveFormat);

            // Set the "ParagraphBreak" to a custom value that we wish to put at the end of every paragraph.
            txtSaveOptions.ParagraphBreak = " End of paragraph.\n\n\t";

            doc.Save(ArtifactsDir + "TxtSaveOptions.ParagraphBreak.txt", txtSaveOptions);

            string docText = File.ReadAllText(ArtifactsDir + "TxtSaveOptions.ParagraphBreak.txt");

            Assert.AreEqual("Paragraph 1. End of paragraph.\n\n\t" +
                            "Paragraph 2. End of paragraph.\n\n\t" +
                            "Paragraph 3. End of paragraph.\n\n\t", docText);
            //ExEnd
        }

        [Test]
        public void Encoding()
        {
            //ExStart
            //ExFor:TxtSaveOptionsBase.Encoding
            //ExSummary:Shows how to set encoding for a .txt output document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some text with characters from outside the ASCII character set.
            builder.Write("À È Ì Ò Ù.");

            // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
            // to modify how we save the document to plaintext.
            TxtSaveOptions txtSaveOptions = new TxtSaveOptions();
            
            // Verify that the "Encoding" property contains the appropriate encoding for our document's contents.
            Assert.AreEqual(System.Text.Encoding.UTF8, txtSaveOptions.Encoding);

            doc.Save(ArtifactsDir + "TxtSaveOptions.Encoding.UTF8.txt", txtSaveOptions);

            string docText = System.Text.Encoding.UTF8.GetString(File.ReadAllBytes(ArtifactsDir + "TxtSaveOptions.Encoding.UTF8.txt"));
            
            Assert.AreEqual("\uFEFFÀ È Ì Ò Ù.\r\n", docText);

            // Using an unsuitable encoding may result in a loss of document contents.
            txtSaveOptions.Encoding = System.Text.Encoding.ASCII;
            doc.Save(ArtifactsDir + "TxtSaveOptions.Encoding.ASCII.txt", txtSaveOptions);
            docText = System.Text.Encoding.ASCII.GetString(File.ReadAllBytes(ArtifactsDir + "TxtSaveOptions.Encoding.ASCII.txt"));

            Assert.AreEqual("? ? ? ? ?.\r\n", docText);
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void PreserveTableLayout(bool preserveTableLayout)
        {
            //ExStart
            //ExFor:TxtSaveOptions.PreserveTableLayout
            //ExSummary:Shows how to preserve the layout of tables when converting to plaintext.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartTable();
            builder.InsertCell();
            builder.Write("Row 1, cell 1");
            builder.InsertCell();
            builder.Write("Row 1, cell 2");
            builder.EndRow();
            builder.InsertCell();
            builder.Write("Row 2, cell 1");
            builder.InsertCell();
            builder.Write("Row 2, cell 2");
            builder.EndTable();

            // Create a "TxtSaveOptions" object, which we can pass to the document's "Save" method
            // to modify how we save the document to plaintext.
            TxtSaveOptions txtSaveOptions = new TxtSaveOptions();

            // Set the "PreserveTableLayout" property to "true" to apply whitespace padding to the contents
            // of the output plaintext document to preserve as much of the table's layout as possible.
            // Set the "PreserveTableLayout" property to "false" to save all tables' contents
            // as a continuous body of text, with just a new line for each row.
            txtSaveOptions.PreserveTableLayout = preserveTableLayout;

            doc.Save(ArtifactsDir + "TxtSaveOptions.PreserveTableLayout.txt", txtSaveOptions);

            string docText = File.ReadAllText(ArtifactsDir + "TxtSaveOptions.PreserveTableLayout.txt");

            if (preserveTableLayout)
                Assert.AreEqual("Row 1, cell 1                Row 1, cell 2\r\n" +
                                "Row 2, cell 1                Row 2, cell 2\r\n\r\n", docText);
            else
                Assert.AreEqual("Row 1, cell 1\r\n" +
                                "Row 1, cell 2\r\n" +
                                "Row 2, cell 1\r\n" +
                                "Row 2, cell 2\r\n\r\n", docText);
            //ExEnd
        }
    }
}