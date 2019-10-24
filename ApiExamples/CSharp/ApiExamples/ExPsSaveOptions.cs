// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExPsSaveOptions : ApiExampleBase
    {
        [Test]
        public void UseBookFoldPrintingSettings()
        {
            //ExStart
            //ExFor:PsSaveOptions
            //ExFor:PsSaveOptions.SaveFormat
            //ExFor:PsSaveOptions.UseBookFoldPrintingSettings
            //ExSummary:Shows how to create a bookfold in the PostScript format.
            Document doc = new Document(MyDir + "Paragraphs.docx");

            // Configure both page setup and PsSaveOptions to create a book fold
            foreach (Section s in doc.Sections)
            {
                s.PageSetup.MultiplePages = MultiplePagesType.BookFoldPrinting;
            }

            PsSaveOptions saveOptions = new PsSaveOptions
            {
                SaveFormat = SaveFormat.Ps,
                UseBookFoldPrintingSettings = true
            };

            // In order to make a booklet, we will need to print this document, stack the pages
            // in the order they come out of the printer and then fold down the middle
            doc.Save(ArtifactsDir + "PsSaveOptions.UseBookFoldPrintingSettings.ps", saveOptions);
            //ExEnd
        }
    }
}