// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Text;
using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    internal class ExHtmlFixedSaveOptions : ApiExampleBase
    {
        [Test]
        public void UseEncoding()
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.Encoding
            //ExSummary:Shows how to set encoding while exporting to HTML.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Writeln("Hello World!");

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                Encoding = new ASCIIEncoding()
            };

            doc.Save(ArtifactsDir + "UseEncoding.html", htmlFixedSaveOptions);
            //ExEnd
        }

        // Note: Test doesn't contain validation result, because it's may take a lot of time for assert result
        // For validation result, you can save the document to HTML file and check out with notepad++, that file encoding will be correctly displayed (Encoding tab in Notepad++)
        [Test]
        public void EncodingUsingGetEncoding()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                Encoding = Encoding.GetEncoding("utf-16")
            };

            doc.Save(ArtifactsDir + "EncodingUsingGetEncoding.html", htmlFixedSaveOptions);
        }

        // Note: Test doesn't contain validation result, because it's may take a lot of time for assert result
        // For validation result, you can save the document to HTML file and check out with notepad++, that file encoding will be correctly displayed (Encoding tab in Notepad++)
        [Test]
        public void ExportEmbeddedObjects()
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.ExportEmbeddedCss
            //ExFor:HtmlFixedSaveOptions.ExportEmbeddedFonts
            //ExFor:HtmlFixedSaveOptions.ExportEmbeddedImages
            //ExFor:HtmlFixedSaveOptions.ExportEmbeddedSvg
            //ExSummary:Shows how to export embedded objects into HTML file.
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                ExportEmbeddedCss = true,
                ExportEmbeddedFonts = true,
                ExportEmbeddedImages = true,
                ExportEmbeddedSvg = true
            };

            doc.Save(ArtifactsDir + "ExportEmbeddedObjects.html", htmlFixedSaveOptions);
            //ExEnd
        }

        [Test]
        public void ExportFormFields()
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.ExportFormFields
            //ExSummary:Show how to exporting form fields from a document into HTML file.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertCheckBox("CheckBox", false, 15);

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                ExportFormFields = true
            };

            doc.Save(ArtifactsDir + "ExportFormFields.html", htmlFixedSaveOptions);
        }

        [Test]
        public void AddCssClassNamesPrefix()
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.CssClassNamesPrefix
            //ExFor:HtmlFixedSaveOptions.SaveFontFaceCssSeparately
            //ExSummary:Shows how to add prefix to all class names in css file.
            Document doc = new Document(MyDir + "Bookmark.doc");

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                CssClassNamesPrefix = "test",
                SaveFontFaceCssSeparately = true
            };

            doc.Save(ArtifactsDir + "cssPrefix_Out.html", htmlFixedSaveOptions);
            //ExEnd

            DocumentHelper.FindTextInFile(ArtifactsDir + "cssPrefix_Out/styles.css", "test");
        }

        [Test]
        public void HorizontalAlignment()
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.PageHorizontalAlignment
            //ExFor:HtmlFixedPageHorizontalAlignment
            //ExSummary:Shows how to set the horizontal alignment of pages in HTML file.
            Document doc = new Document(MyDir + "Bookmark.doc");

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                PageHorizontalAlignment = HtmlFixedPageHorizontalAlignment.Left
            };

            doc.Save(ArtifactsDir + "HtmlFixedPageHorizontalAlignment.html", htmlFixedSaveOptions);
            //ExEnd
        }

        [Test]
        public void PageMargins()
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.PageMargins
            //ExSummary:Shows how to set the margins around pages in HTML file.
            Document doc = new Document(MyDir + "Bookmark.doc");

            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions
            {
                PageMargins = 10
            };

            doc.Save(ArtifactsDir + "HtmlFixedPageMargins.html", saveOptions);
        }

        [Test]
        public void PageMarginsException()
        {
            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
            Assert.That(() => saveOptions.PageMargins = -1, Throws.TypeOf<ArgumentException>());
        }

        [Test]
        public void OptimizeGraphicsOutput()
        {
            //ExStart
            //ExFor:FixedPageSaveOptions.OptimizeOutput
            //ExSummary:Shows how to optimize document objects while saving to html.
            Document doc = new Document(MyDir + "HtmlFixedSaveOptions.OptimizeGraphicsOutput.doc");

            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions { OptimizeOutput = false };

            doc.Save(ArtifactsDir + "HtmlFixedPageMargins.html", saveOptions);
            //ExEnd
        }

        //ExStart
        //ExFor:HtmlFixedSaveOptions.UseTargetMachineFonts
        //ExFor:IResourceSavingCallback
        //ExFor:IResourceSavingCallback.ResourceSaving(ResourceSavingArgs)
        //ExSummary: Shows how used target machine fonts to display the document.
        [Test] //ExSkip
        public void UsingMachineFonts()
        {
            Document doc = new Document(MyDir + "Font.DisappearingBulletPoints.doc");

            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions
            {
                UseTargetMachineFonts = true,
                FontFormat = ExportFontFormat.Ttf,
                ExportEmbeddedFonts = false,
                ResourceSavingCallback = new ResourceSavingCallback()
            };

            doc.Save(ArtifactsDir + "UseMachineFonts.html", saveOptions);
        }

        private class ResourceSavingCallback : IResourceSavingCallback
        {
            /// <summary>
            /// Called when Aspose.Words saves an external resource to fixed page HTML or SVG.
            /// </summary>
            public void ResourceSaving(ResourceSavingArgs args)
            {
                args.ResourceStream = new MemoryStream();
                args.KeepResourceStreamOpen = true;

                string extension = Path.GetExtension(args.ResourceFileName);
                switch (extension)
                {
                    case ".ttf":
                    case ".woff":
                    {
                        Assert.Fail(
                            "'ResourceSavingCallback' is not fired for fonts when 'UseTargetMachineFonts' is true");
                        break;
                    }
                }
            }
        }

        //ExEnd
    }
}