// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
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
        [Test]
        public void PageBreaks()
        {
            //ExStart
            //ExFor:TxtSaveOptionsBase.ForcePageBreaks
            //ExSummary:Shows how to specify whether the page breaks should be preserved during export.
            Document doc = new Document(MyDir + "SaveOptions.PageBreaks.docx");

            TxtSaveOptions saveOptions = new TxtSaveOptions { ForcePageBreaks = false };

            doc.Save(ArtifactsDir + "SaveOptions.PageBreaks.txt", saveOptions);
            //ExEnd
        }

        [Test]
        public void AddBidiMarks()
        {
            //ExStart
            //ExFor:TxtSaveOptions.AddBidiMarks
            //ExSummary:Shows how to insert Unicode Character 'RIGHT-TO-LEFT MARK' (U+200F) before each bi-directional Run in text.
            Document doc = new Document(MyDir + "Document.docx");
            
            TxtSaveOptions saveOptions = new TxtSaveOptions { AddBidiMarks = true };

            doc.Save(ArtifactsDir + "AddBidiMarks.txt", saveOptions);
            //ExEnd
        }

        [Test]
        [TestCase(TxtExportHeadersFootersMode.None)]
        [TestCase(TxtExportHeadersFootersMode.AllAtEnd)]
        [TestCase(TxtExportHeadersFootersMode.PrimaryOnly)]
        public void ExportHeadersFooters(TxtExportHeadersFootersMode txtExportHeadersFootersMode)
        {
            //ExStart
            //ExFor:TxtSaveOptionsBase.ExportHeadersFootersMode
            //ExFor:TxtExportHeadersFootersMode
            //ExSummary:Shows how to specifies the way headers and footers are exported to plain text format.
            Document doc = new Document(MyDir + "HeaderFooter.HeaderFooterOrder.docx");

            // Three values are available in TxtExportHeadersFootersMode enum:
            // "None" - No headers and footers are exported
            // "AllAtEnd" - All headers and footers are placed after all section bodies at the very end of a document
            // "PrimaryOnly" - Only primary headers and footers are exported at the beginning and end of each section (default value)
            TxtSaveOptions saveOptions = new TxtSaveOptions { ExportHeadersFootersMode = txtExportHeadersFootersMode };

            doc.Save(ArtifactsDir + "ExportHeadersFooters.txt", saveOptions);
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
            Document doc = new Document(MyDir + "TxtSaveOptions.TxtListIndentation.docx");
 
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
            TxtSaveOptions txtSaveOptions = new TxtSaveOptions { ParagraphBreak = " End of paragraph.\n\n\t" };
            
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
            //ExFor:TxtSaveOptionsBase.PreserveTableLayout
            //ExFor:TxtSaveOptionsBase.SimplifyListLabels
            //ExSummary:Shows how to change the appearance of tables and lists during conversion to a txt document output.
            // Open a document with a table
            Document doc = new Document(MyDir + "Lists.PrintOutAllLists.doc");

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