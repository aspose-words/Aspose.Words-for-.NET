// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExXpsSaveOptions : ApiExampleBase
    {
        [Test]
        public void OutlineLevels()
        {
            //ExStart
            //ExFor:XpsSaveOptions
            //ExFor:XpsSaveOptions.#ctor
            //ExFor:XpsSaveOptions.OutlineOptions
            //ExFor:XpsSaveOptions.SaveFormat
            //ExSummary:Shows how to limit the headings' level that will appear in the outline of a saved XPS document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert headings that can serve as TOC entries of levels 1, 2, and then 3.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            Assert.True(builder.ParagraphFormat.IsHeading);

            builder.Writeln("Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

            builder.Writeln("Heading 1.1");
            builder.Writeln("Heading 1.2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;

            builder.Writeln("Heading 1.2.1");
            builder.Writeln("Heading 1.2.2");

            // Create an "XpsSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .XPS.
            XpsSaveOptions saveOptions = new XpsSaveOptions();

            Assert.AreEqual(SaveFormat.Xps, saveOptions.SaveFormat);

            // The output XPS document will contain an outline, a table of contents that lists headings in the document body.
            // Clicking on an entry in this outline will take us to the location of its respective heading.
            // Set the "HeadingsOutlineLevels" property to "2" to exclude all headings whose levels are above 2 from the outline.
            // The last two headings we have inserted above will not appear.
            saveOptions.OutlineOptions.HeadingsOutlineLevels = 2;

            doc.Save(ArtifactsDir + "XpsSaveOptions.OutlineLevels.xps", saveOptions);
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void BookFold(bool renderTextAsBookFold)
        {
            //ExStart
            //ExFor:XpsSaveOptions.#ctor(SaveFormat)
            //ExFor:XpsSaveOptions.UseBookFoldPrintingSettings
            //ExSummary:Shows how to save a document to the XPS format in the form of a book fold.
            Document doc = new Document(MyDir + "Paragraphs.docx");

            // Create an "XpsSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .XPS.
            XpsSaveOptions xpsOptions = new XpsSaveOptions(SaveFormat.Xps);

            // Set the "UseBookFoldPrintingSettings" property to "true" to arrange the contents
            // in the output XPS in a way that helps us use it to make a booklet.
            // Set the "UseBookFoldPrintingSettings" property to "false" to render the XPS normally.
            xpsOptions.UseBookFoldPrintingSettings = renderTextAsBookFold;

            // If we are rendering the document as a booklet, we must set the "MultiplePages"
            // properties of the page setup objects of all sections to "MultiplePagesType.BookFoldPrinting".
            if (renderTextAsBookFold)
                foreach (Section s in doc.Sections)
                {
                    s.PageSetup.MultiplePages = MultiplePagesType.BookFoldPrinting;
                }

            // Once we print this document, we can turn it into a booklet by stacking the pages
            // to come out of the printer and folding down the middle.
            doc.Save(ArtifactsDir + "XpsSaveOptions.BookFold.xps", xpsOptions);
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void OptimizeOutput(bool optimizeOutput)
        {
            //ExStart
            //ExFor:FixedPageSaveOptions.OptimizeOutput
            //ExSummary:Shows how to optimize document objects while saving to xps.
            Document doc = new Document(MyDir + "Unoptimized document.docx");

            // Create an "XpsSaveOptions" object to pass to the document's "Save" method
            // to modify how that method converts the document to .XPS.
            XpsSaveOptions saveOptions = new XpsSaveOptions();

            // Set the "OptimizeOutput" property to "true" to take measures such as removing nested or empty canvases
            // and concatenating adjacent runs with identical formatting to optimize the output document's content.
            // This may affect the appearance of the document.
            // Set the "OptimizeOutput" property to "false" to save the document normally.
            saveOptions.OptimizeOutput = optimizeOutput;

            doc.Save(ArtifactsDir + "XpsSaveOptions.OptimizeOutput.xps", saveOptions);
            //ExEnd

            FileInfo outFileInfo = new FileInfo(ArtifactsDir + "XpsSaveOptions.OptimizeOutput.xps");

            if (optimizeOutput)
                Assert.That(50000, Is.AtLeast(outFileInfo.Length));
            else
                Assert.That(60000, Is.LessThan(outFileInfo.Length));

            TestUtil.DocPackageFileContainsString(
                optimizeOutput
                    ? "Glyphs OriginX=\"34.294998169\" OriginY=\"10.31799984\" " +
                      "UnicodeString=\"This document contains complex content which can be optimized to save space when \""
                    : "<Glyphs OriginX=\"34.294998169\" OriginY=\"10.31799984\" UnicodeString=\"This\"",
                ArtifactsDir + "XpsSaveOptions.OptimizeOutput.xps", "1.fpage");
        }

        [Test]
        public void ExportExactPages()
        {
            //ExStart
            //ExFor:FixedPageSaveOptions.PageSet
            //ExFor:PageSet.#ctor(int[])
            //ExSummary:Shows how to extract pages based on exact page indices.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add five pages to the document.
            for (int i = 1; i < 6; i++)
            {
                builder.Write("Page " + i);
                builder.InsertBreak(BreakType.PageBreak);
            }

            // Create an "XpsSaveOptions" object, which we can pass to the document's "Save" method
            // to modify how that method converts the document to .XPS.
            XpsSaveOptions xpsOptions = new XpsSaveOptions();

            // Use the "PageSet" property to select a set of the document's pages to save to output XPS.
            // In this case, we will choose, via a zero-based index, only three pages: page 1, page 2, and page 4.
            xpsOptions.PageSet = new PageSet(0, 1, 3);

            doc.Save(ArtifactsDir + "XpsSaveOptions.ExportExactPages.xps", xpsOptions);
            //ExEnd
        }
    }
}