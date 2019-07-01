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
            //ExSummary:Shows how list levels are displayed when the document is converting to plain text format
            Document doc = new Document(MyDir + "TxtSaveOptions.TxtListIndentation.docx");
 
            TxtSaveOptions txtSaveOptions = new TxtSaveOptions();
            txtSaveOptions.ListIndentation.Count = 3;
            txtSaveOptions.ListIndentation.Character = ' ';
            txtSaveOptions.PreserveTableLayout = true;
 
            doc.Save(ArtifactsDir + "TxtSaveOptions.TxtListIndentation.txt", txtSaveOptions);
            //ExEnd
        }
    }
}