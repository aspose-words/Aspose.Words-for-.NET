// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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
        public void SetZoom()
        {
            //ExStart
            //ExFor:Document.ViewOptions
            //ExFor:ViewOptions
            //ExFor:ViewOptions.ViewType
            //ExFor:ViewOptions.ZoomType
            //ExFor:ViewOptions.ZoomPercent
            //ExFor:ViewType
            //ExSummary:Shows how to make sure the document is displayed at 50% zoom when opened in Microsoft Word.
            Document doc = new Document(MyDir + "Document.docx");

            // We can set the zoom factor to a percentage
            doc.ViewOptions.ViewType = ViewType.PageLayout;
            doc.ViewOptions.ZoomPercent = 50;

            // Or we can set the ZoomType to a different value to avoid using percentages 
            Assert.AreEqual(ZoomType.None, doc.ViewOptions.ZoomType);

            doc.Save(ArtifactsDir + "ViewOptions.SetZoom.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ViewOptions.SetZoom.docx");

            Assert.AreEqual(ViewType.PageLayout, doc.ViewOptions.ViewType);
            Assert.AreEqual(50.0d, doc.ViewOptions.ZoomPercent);
            Assert.AreEqual(ZoomType.None, doc.ViewOptions.ZoomType);
        }

        [Test]
        public void DisplayBackgroundShape()
        {
            //ExStart
            //ExFor:ViewOptions.DisplayBackgroundShape
            //ExSummary:Shows how to hide/display document background images in view options.
            // Create a new document from an html string with a flat background color
            const string html = @"<html>
                <body style='background-color: blue'>
                    <p>Hello world!</p>
                </body>
            </html>";

            Document doc = new Document(new MemoryStream(Encoding.Unicode.GetBytes(html)));

            // The source for the document has a flat color background, the presence of which will turn on the DisplayBackgroundShape flag
            // We can disable it like this
            doc.ViewOptions.DisplayBackgroundShape = false;

            doc.Save(ArtifactsDir + "ViewOptions.DisplayBackgroundShape.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ViewOptions.DisplayBackgroundShape.docx");

            Assert.False(doc.ViewOptions.DisplayBackgroundShape);
        }

        [Test]
        public void DisplayPageBoundaries()
        {
            //ExStart
            //ExFor:ViewOptions.DoNotDisplayPageBoundaries
            //ExSummary:Shows how to hide vertical whitespace and headers/footers in view options.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert content spanning 3 pages
            builder.Writeln("Paragraph 1, Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Paragraph 2, Page 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Paragraph 3, Page 3");

            // Insert a header and a footer
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Writeln("Header");
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Writeln("Footer");

            // In this case we have a lot of space taken up by quite a little amount of content
            // In older versions of Microsoft Word, we can hide headers/footers and compact vertical whitespace of pages
            // to give the document's main body content some flow by setting this flag
            doc.ViewOptions.DoNotDisplayPageBoundaries = true;

            doc.Save(ArtifactsDir + "ViewOptions.DisplayPageBoundaries.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "ViewOptions.DisplayPageBoundaries.docx");

            Assert.True(doc.ViewOptions.DoNotDisplayPageBoundaries);
        }

        [TestCase(false)]
        [TestCase(true)]
        public void FormsDesign(bool useFormsDesign)
        {
            //ExStart
            //ExFor:ViewOptions.FormsDesign
            //ExFor:WordML2003SaveOptions
            //ExFor:WordML2003SaveOptions.SaveFormat
            //ExSummary:Shows how to save to a .wml document while applying save options.
            Document doc = new Document(MyDir + "Document.docx");

            WordML2003SaveOptions options = new WordML2003SaveOptions()
            {
                SaveFormat = SaveFormat.WordML,
                MemoryOptimization = true,
                PrettyFormat = true
            };

            // Enables forms design mode in WordML documents
            doc.ViewOptions.FormsDesign = useFormsDesign;

            doc.Save(ArtifactsDir + "ViewOptions.FormsDesign.xml", options);

            Assert.AreEqual(useFormsDesign,
                File.ReadAllText(ArtifactsDir + "ViewOptions.FormsDesign.xml").Contains("<w:formsDesign />"));
            //ExEnd
        }
    }
}