// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
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
            //ExFor:TxtSaveOptions.ForcePageBreaks
            //ExSummary:Shows how to specify whether the page breaks should be preserved during export.
            Document doc = new Document(MyDir + "SaveOptions.PageBreaks.docx");

            TxtSaveOptions saveOptions = new TxtSaveOptions { ForcePageBreaks = false };

            doc.Save(MyDir + @"\Artifacts\SaveOptions.PageBreaks.txt", saveOptions);
            //ExEnd
        }

        [Test]
        public void AddBidiMarks()
        {
            //ExStart
            //ExFor:TxtSaveOptions.AddBidiMarks
            //ExSummary:Shows how to insert Unicode Character 'RIGHT-TO-LEFT MARK' (U+200F) before each bi-directional Run in text.
            Document doc = new Document(MyDir + "Document.docx");
            // In Aspose.Words by default this option is set to true unlike Word
            TxtSaveOptions saveOptions = new TxtSaveOptions { AddBidiMarks = false };

            doc.Save(MyDir + @"\Artifacts\AddBidiMarks.txt", saveOptions);
            //ExEnd
        }

        [Test]
        [TestCase(TxtExportHeadersFootersMode.None)]
        [TestCase(TxtExportHeadersFootersMode.AllAtEnd)]
        [TestCase(TxtExportHeadersFootersMode.PrimaryOnly)]
        public void ExportHeadersFooters(TxtExportHeadersFootersMode txtExportHeadersFootersMode)
        {
            //ExStart
            //ExFor:TxtSaveOptions.ExportHeadersFootersMode
            //ExFor:TxtExportHeadersFootersMode
            //ExSummary:Shows how to specifies the way headers and footers are exported to plain text format.
            Document doc = new Document(MyDir + "HeaderFooter.HeaderFooterOrder.docx");

            // Three values are available in TxtExportHeadersFootersMode enum:
            // "None" - No headers and footers are exported
            // "AllAtEnd" - All headers and footers are placed after all section bodies at the very end of a document
            // "PrimaryOnly" - Only primary headers and footers are exported at the beginning and end of each section (default value)
            TxtSaveOptions saveOptions = new TxtSaveOptions { ExportHeadersFootersMode = txtExportHeadersFootersMode };

            doc.Save(MyDir + @"\Artifacts\ExportHeadersFooters.txt", saveOptions);
            //ExEnd
        }
    }
}