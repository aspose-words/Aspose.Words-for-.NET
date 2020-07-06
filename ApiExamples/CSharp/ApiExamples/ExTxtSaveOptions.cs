// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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
            //ExSummary:Shows how to specify whether the page breaks should be preserved during export.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 3");

            // If ForcePageBreaks is set to true then the output document will have form feed characters in place of page breaks
            // Otherwise, they will be line breaks
            TxtSaveOptions saveOptions = new TxtSaveOptions { ForcePageBreaks = forcePageBreaks };

            doc.Save(ArtifactsDir + "TxtSaveOptions.PageBreaks.txt", saveOptions);
            
            // If we load the document using Aspose.Words again, the page breaks will be preserved/lost depending on ForcePageBreaks
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

            TxtSaveOptions saveOptions = new TxtSaveOptions { AddBidiMarks = addBidiMarks, Encoding = System.Text.Encoding.Unicode};

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
            //ExSummary:Shows how to specifies the way headers and footers are exported to plain text format.
            Document doc = new Document();

            // Insert even and primary headers/footers into the document
            // The primary header/footers should override the even ones 
            doc.FirstSection.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.HeaderEven));
            doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderEven].AppendParagraph("Even header");
            doc.FirstSection.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.FooterEven));
            doc.FirstSection.HeadersFooters[HeaderFooterType.FooterEven].AppendParagraph("Even footer");
            doc.FirstSection.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.HeaderPrimary));
            doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].AppendParagraph("Primary header");
            doc.FirstSection.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.FooterPrimary));
            doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary].AppendParagraph("Primary footer");

            // Insert pages that would display these headers and footers
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2");
            builder.InsertBreak(BreakType.PageBreak); 
            builder.Write("Page 3");

            // Three values are available in TxtExportHeadersFootersMode enum:
            // "None" - No headers and footers are exported
            // "AllAtEnd" - All headers and footers are placed after all section bodies at the very end of a document
            // "PrimaryOnly" - Only primary headers and footers are exported at the beginning and end of each section (default value)
            TxtSaveOptions saveOptions = new TxtSaveOptions { ExportHeadersFootersMode = txtExportHeadersFootersMode };
            
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
            //ExSummary:Shows how to configure list indenting when converting to plaintext.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a list with three levels of indentation
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Item 1");
            builder.ListFormat.ListIndent();
            builder.Writeln("Item 2");
            builder.ListFormat.ListIndent(); 
            builder.Write("Item 3");

            // Microsoft Word list objects get lost when converting to plaintext
            // We can create a custom representation for list indentation using pure plaintext with a SaveOptions object
            // In this case, each list item will be left-padded by 3 space characters times its list indent level
            TxtSaveOptions txtSaveOptions = new TxtSaveOptions();
            txtSaveOptions.ListIndentation.Count = 3;
            txtSaveOptions.ListIndentation.Character = ' ';

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
            //ExSummary:Shows how to change the appearance of lists when converting to plaintext.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a bulleted list with five levels of indentation
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

            // The SimplifyListLabels flag will convert some list symbols
            // into ASCII characters such as *, o, +, > etc, depending on list level
            TxtSaveOptions txtSaveOptions = new TxtSaveOptions { SimplifyListLabels = simplifyListLabels };

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
            // Create a new document and add some paragraphs
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Paragraph 1.");
            builder.Writeln("Paragraph 2.");
            builder.Write("Paragraph 3.");

            // When saved to plain text, the paragraphs we created can be separated by a custom string
            TxtSaveOptions txtSaveOptions = new TxtSaveOptions { SaveFormat = SaveFormat.Text, ParagraphBreak = " End of paragraph.\n\n\t" };
            
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
            // Create a new document and add some text from outside the ASCII character set
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("À È Ì Ò Ù.");

            // We can use a SaveOptions object to make sure the encoding we save the .txt document in supports our content
            TxtSaveOptions txtSaveOptions = new TxtSaveOptions { Encoding = System.Text.Encoding.UTF8 };

            doc.Save(ArtifactsDir + "TxtSaveOptions.Encoding.txt", txtSaveOptions);

            string docText = System.Text.Encoding.UTF8.GetString(File.ReadAllBytes(ArtifactsDir + "TxtSaveOptions.Encoding.txt"));
            
            Assert.AreEqual("\uFEFFÀ È Ì Ò Ù.\r\n", docText);
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void TableLayout(bool preserveTableLayout)
        {
            //ExStart
            //ExFor:TxtSaveOptions.PreserveTableLayout
            //ExSummary:Shows how to preserve the layout of tables when converting to plaintext.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table
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

            // Tables, with their borders and widths do not translate to plaintext
            // However, we can configure a SaveOptions object to arrange table contents to preserve some of the table's appearance
            TxtSaveOptions txtSaveOptions = new TxtSaveOptions { PreserveTableLayout = preserveTableLayout };

            doc.Save(ArtifactsDir + "TxtSaveOptions.TableLayout.txt", txtSaveOptions);

            string docText = File.ReadAllText(ArtifactsDir + "TxtSaveOptions.TableLayout.txt");

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