// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
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
    internal class ExXlsxSaveOptions : ApiExampleBase
    {
        [Test]
        public void CompressXlsx()
        {
            //ExStart
            //ExFor:XlsxSaveOptions.CompressionLevel
            //ExSummary:Shows how to compress XLSX document.
            Document doc = new Document(MyDir + "Shape with linked chart.docx");

            XlsxSaveOptions xlsxSaveOptions = new XlsxSaveOptions();
            xlsxSaveOptions.CompressionLevel = CompressionLevel.Maximum; 

            doc.Save(ArtifactsDir + "XlsxSaveOptions.CompressXlsx.xlsx", xlsxSaveOptions);
            //ExEnd
        }

        [Test]
        public void SelectionMode()
        {
            //ExStart:SelectionMode
            //GistId:470c0da51e4317baae82ad9495747fed
            //ExFor:XlsxSaveOptions.SectionMode
            //ExSummary:Shows how to save document as a separate worksheets.
            Document doc = new Document(MyDir + "Big document.docx");

            // Each section of a document will be created as a separate worksheet.
            // Use 'SingleWorksheet' to display all document on one worksheet.
            XlsxSaveOptions xlsxSaveOptions = new XlsxSaveOptions();
            xlsxSaveOptions.SectionMode = XlsxSectionMode.MultipleWorksheets;

            doc.Save(ArtifactsDir + "XlsxSaveOptions.SelectionMode.xlsx", xlsxSaveOptions);
            //ExEnd:SelectionMode
        }

        [Test]
        public void DateTimeParsingMode()
        {
            //ExStart:DateTimeParsingMode
            //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
            //ExFor:XlsxSaveOptions.DateTimeParsingMode
            //ExFor:XlsxDateTimeParsingMode
            //ExSummary:Shows how to specify autodetection of the date time format.
            Document doc = new Document(MyDir + "Xlsx DateTime.docx");

            XlsxSaveOptions saveOptions = new XlsxSaveOptions();
            // Specify using datetime format autodetection.
            saveOptions.DateTimeParsingMode = XlsxDateTimeParsingMode.Auto;

            doc.Save(ArtifactsDir + "XlsxSaveOptions.DateTimeParsingMode.xlsx", saveOptions);
            //ExEnd:DateTimeParsingMode
        }
    }
}
