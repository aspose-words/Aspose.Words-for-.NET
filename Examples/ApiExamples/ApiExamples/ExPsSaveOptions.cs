// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
        [TestCase(false)]
        [TestCase(true)]
        public void UseBookFoldPrintingSettings(bool renderTextAsBookFold)
        {
            //ExStart
            //ExFor:PsSaveOptions
            //ExFor:PsSaveOptions.SaveFormat
            //ExFor:PsSaveOptions.UseBookFoldPrintingSettings
            //ExSummary:Shows how to save a document to the Postscript format in the form of a book fold.
            Document doc = new Document(MyDir + "Paragraphs.docx");

            // Create a "PsSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to PostScript.
            // Set the "UseBookFoldPrintingSettings" property to "true" to arrange the contents
            // in the output Postscript document in a way that helps us make a booklet out of it.
            // Set the "UseBookFoldPrintingSettings" property to "false" to save the document normally.
            PsSaveOptions saveOptions = new PsSaveOptions
            {
                SaveFormat = SaveFormat.Ps,
                UseBookFoldPrintingSettings = renderTextAsBookFold
            };

            // If we are rendering the document as a booklet, we must set the "MultiplePages"
            // properties of the page setup objects of all sections to "MultiplePagesType.BookFoldPrinting".
            foreach (Section s in doc.Sections)
            {
                s.PageSetup.MultiplePages = MultiplePagesType.BookFoldPrinting;
            }

            // Once we print this document on both sides of the pages, we can fold all the pages down the middle at once,
            // and the contents will line up in a way that creates a booklet.
            doc.Save(ArtifactsDir + "PsSaveOptions.UseBookFoldPrintingSettings.ps", saveOptions);
            //ExEnd
        }
    }
}