// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExViewOptions : ApiExampleBase
    {
        [Test]
        public void SetZoomPercentage()
        {
            //ExStart
            //ExFor:Document.ViewOptions
            //ExFor:ViewOptions
            //ExFor:ViewOptions.ViewType
            //ExFor:ViewOptions.ZoomPercent
            //ExFor:ViewOptions.ZoomType
            //ExFor:ViewType
            //ExSummary:Shows how to set a custom zoom factor, which older versions of Microsoft Word will apply to a document upon loading.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            doc.ViewOptions.ViewType = ViewType.PageLayout;
            doc.ViewOptions.ZoomPercent = 50;

            Assert.AreEqual(ZoomType.Custom, doc.ViewOptions.ZoomType);
            Assert.AreEqual(ZoomType.None, doc.ViewOptions.ZoomType);

            doc.Save(ArtifactsDir + "ViewOptions.SetZoomPercentage.doc");
            //ExEnd

            doc = new Document(ArtifactsDir + "ViewOptions.SetZoomPercentage.doc");

            Assert.AreEqual(ViewType.PageLayout, doc.ViewOptions.ViewType);
            Assert.AreEqual(50.0d, doc.ViewOptions.ZoomPercent);
            Assert.AreEqual(ZoomType.None, doc.ViewOptions.ZoomType);
        }

        [TestCase(ZoomType.PageWidth)]
        [TestCase(ZoomType.FullPage)]
        [TestCase(ZoomType.TextFit)]
        public void SetZoomType(ZoomType zoomType)
        {
            //ExStart
            //ExFor:Document.ViewOptions
            //ExFor:ViewOptions
            //ExFor:ViewOptions.ZoomType
            //ExSummary:Shows how to set a custom zoom type, which older versions of Microsoft Word will apply to a document upon loading.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // Set the "ZoomType" property to "ZoomType.PageWidth" to get Microsoft Word
            // to automatically zoom the document to fit the width of the page.
            // Set the "ZoomType" property to "ZoomType.FullPage" to get Microsoft Word
            // to automatically zoom the document to make the entire first page visible.
            // Set the "ZoomType" property to "ZoomType.TextFit" to get Microsoft Word
            // to automatically zoom the document to fit the inner text margins of the first page.
            doc.ViewOptions.ZoomType = zoomType;

            doc.Save(ArtifactsDir + "ViewOptions.SetZoomType.doc");
            //ExEnd

            doc = new Document(ArtifactsDir + "ViewOptions.SetZoomType.doc");

            Assert.AreEqual(zoomType, doc.ViewOptions.ZoomType);
        }

        [TestCase(false)]
        [TestCase(true)]
        public void DisplayBackgroundShape(bool displayBackgroundShape)
        {
            //ExStart
            //ExFor:ViewOptions.DisplayBackgroundShape
            //ExSummary:Shows how to hide/display document background images in view options.
            // Use an HTML string to create a new document with a flat background color.
            const string html = 
            @"<html>
                <body style='background-color: blue'>
                    <p>Hello world!</p>
                </body>
            </html>";

            Document doc = new Document(new MemoryStream(Encoding.Unicode.GetBytes(html)));

            // The source for the document has a flat color background,
            // the presence of which will set the "DisplayBackgroundShape" flag to "true".
            Assert.True(doc.ViewOptions.DisplayBackgroundShape);

            // Keep the "DisplayBackgroundShape" as "true" to get the document to display the background color.
            // This may affect some text colors to improve visibility.
            // Set the "DisplayBackgroundShape" to "false" to not display the background color.
            doc.ViewOptions.DisplayBackgroundShape = displayBackgroundShape;

            doc.Save(ArtifactsDir + "ViewOptions.DisplayBackgroundShape.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ViewOptions.DisplayBackgroundShape.docx");

            Assert.AreEqual(displayBackgroundShape, doc.ViewOptions.DisplayBackgroundShape);
        }

        [TestCase(false)]
        [TestCase(true)]
        public void DisplayPageBoundaries(bool doNotDisplayPageBoundaries)
        {
            //ExStart
            //ExFor:ViewOptions.DoNotDisplayPageBoundaries
            //ExSummary:Shows how to hide vertical whitespace and headers/footers in view options.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert content that spans across 3 pages.
            builder.Writeln("Paragraph 1, Page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Paragraph 2, Page 2.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Paragraph 3, Page 3.");

            // Insert a header and a footer.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("This is the header.");
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Writeln("This is the footer.");

            // This document contains a small amount of content that takes up a few full pages worth of space.
            // Set the "DoNotDisplayPageBoundaries" flag to "true" to get older versions of Microsoft Word to omit headers,
            // footers, and much of the vertical whitespace when displaying our document.
            // Set the "DoNotDisplayPageBoundaries" flag to "false" to get older versions of Microsoft Word
            // to normally display our document.
            doc.ViewOptions.DoNotDisplayPageBoundaries = doNotDisplayPageBoundaries;

            doc.Save(ArtifactsDir + "ViewOptions.DisplayPageBoundaries.doc");
            //ExEnd

            doc = new Document(ArtifactsDir + "ViewOptions.DisplayPageBoundaries.doc");

            Assert.AreEqual(doNotDisplayPageBoundaries, doc.ViewOptions.DoNotDisplayPageBoundaries);
        }

        [TestCase(false)]
        [TestCase(true)]
        public void FormsDesign(bool useFormsDesign)
        {
            //ExStart
            //ExFor:ViewOptions.FormsDesign
            //ExSummary:Shows how to enable/disable forms design mode.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // Set the "FormsDesign" property to "false" to keep forms design mode disabled.
            // Set the "FormsDesign" property to "true" to enable forms design mode.
            doc.ViewOptions.FormsDesign = useFormsDesign;

            doc.Save(ArtifactsDir + "ViewOptions.FormsDesign.xml");

            Assert.AreEqual(useFormsDesign,
                File.ReadAllText(ArtifactsDir + "ViewOptions.FormsDesign.xml").Contains("<w:formsDesign />"));
            //ExEnd
        }
    }
}