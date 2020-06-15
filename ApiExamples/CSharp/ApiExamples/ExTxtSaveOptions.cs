// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

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

        [Test]
        public void AddBidiMarks()
        {
            //ExStart
            //ExFor:TxtSaveOptions.AddBidiMarks
            //ExSummary:Shows how to insert Unicode Character 'RIGHT-TO-LEFT MARK' (U+200F) before each bi-directional Run in text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ParagraphFormat.Bidi = true;
            builder.Writeln("שלום");

            TxtSaveOptions saveOptions = new TxtSaveOptions { AddBidiMarks = true };

            doc.Save(ArtifactsDir + "TxtSaveOptions.AddBidiMarks.txt", saveOptions);
            //ExEnd
        }

        [TestCase(TxtExportHeadersFootersMode.None)]
        [TestCase(TxtExportHeadersFootersMode.AllAtEnd)]
        [TestCase(TxtExportHeadersFootersMode.PrimaryOnly)]
        public void ExportHeadersFooters(TxtExportHeadersFootersMode txtExportHeadersFootersMode)
        {
            //ExStart
            //ExFor:TxtSaveOptionsBase.ExportHeadersFootersMode
            //ExFor:TxtExportHeadersFootersMode
            //ExSummary:Shows how to specifies the way headers and footers are exported to plain text format.
            Document doc = new Document(MyDir + "Header and footer types.docx");

            // Three values are available in TxtExportHeadersFootersMode enum:
            // "None" - No headers and footers are exported
            // "AllAtEnd" - All headers and footers are placed after all section bodies at the very end of a document
            // "PrimaryOnly" - Only primary headers and footers are exported at the beginning and end of each section (default value)
            TxtSaveOptions saveOptions = new TxtSaveOptions { ExportHeadersFootersMode = txtExportHeadersFootersMode };

            doc.Save(ArtifactsDir + "TxtSaveOptions.ExportHeadersFooters.txt", saveOptions);
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
            //ExSummary:Shows how list levels are displayed when the document is converting to plain text format.
            Document doc = new Document(MyDir + "List indentation.docx");
 
            TxtSaveOptions txtSaveOptions = new TxtSaveOptions();
            txtSaveOptions.ListIndentation.Count = 3;
            txtSaveOptions.ListIndentation.Character = ' ';
            txtSaveOptions.PreserveTableLayout = true;

            doc.Save(ArtifactsDir + "TxtSaveOptions.TxtListIndentation.txt", txtSaveOptions);
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
            builder.Writeln("À È Ì Ò Ù.");

            // We can use a SaveOptions object to make sure the encoding we save the .txt document in supports our content
            TxtSaveOptions txtSaveOptions = new TxtSaveOptions { Encoding = System.Text.Encoding.UTF8 };

            doc.Save(ArtifactsDir + "TxtSaveOptions.Encoding.txt", txtSaveOptions);
            //ExEnd
        }

        [Test]
        public void Appearance()
        {
            //ExStart
            //ExFor:TxtSaveOptions.PreserveTableLayout
            //ExFor:TxtSaveOptions.SimplifyListLabels
            //ExSummary:Shows how to change the appearance of tables and lists during conversion to a txt document output.
            // Open a document with a table
            Document doc = new Document(MyDir + "Rendering.docx");

            // Due to the nature of text documents, table grids and text wrapping will be lost during conversion
            // from a file type that supports tables
            // We can preserve some of the table layout in the appearance of our content with the PreserveTableLayout flag
            // The SimplifyListLabels flag will convert some list symbols
            // into ASCII characters such as *, o, +, > etc, depending on list level
            TxtSaveOptions txtSaveOptions = new TxtSaveOptions { SimplifyListLabels = true, PreserveTableLayout = true};

            doc.Save(ArtifactsDir + "TxtSaveOptions.Appearance.txt", txtSaveOptions);
            //ExEnd
        }
    }
}