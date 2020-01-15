// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.UseEncoding.html", htmlFixedSaveOptions);
            //ExEnd
        }

        // Note: Test doesn't contain validation result, because it's may take a lot of time for assert result
        // For validation result, you can save the document to HTML file and check out with notepad++, that file encoding will be correctly displayed (Encoding tab in Notepad++)
        [Test]
        public void GetEncoding()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                Encoding = Encoding.GetEncoding("utf-16")
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.GetEncoding.html", htmlFixedSaveOptions);
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

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedObjects.html", htmlFixedSaveOptions);
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

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.ExportFormFields.html", htmlFixedSaveOptions);
            //ExEnd
        }

        [Test]
        public void AddCssClassNamesPrefix()
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.CssClassNamesPrefix
            //ExFor:HtmlFixedSaveOptions.SaveFontFaceCssSeparately
            //ExSummary:Shows how to add prefix to all class names in css file.
            Document doc = new Document(MyDir + "Bookmarks.docx");

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                CssClassNamesPrefix = "test",
                SaveFontFaceCssSeparately = true
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.AddCssClassNamesPrefix.html", htmlFixedSaveOptions);
            //ExEnd

            DocumentHelper.FindTextInFile(ArtifactsDir + "HtmlFixedSaveOptions.AddCssClassNamesPrefix/styles.css", "test");
        }

        [Test]
        public void HorizontalAlignment()
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.PageHorizontalAlignment
            //ExFor:HtmlFixedPageHorizontalAlignment
            //ExSummary:Shows how to set the horizontal alignment of pages in HTML file.
            Document doc = new Document(MyDir + "Bookmarks.docx");

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                PageHorizontalAlignment = HtmlFixedPageHorizontalAlignment.Left
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.HorizontalAlignment.html", htmlFixedSaveOptions);
            //ExEnd
        }

        [Test]
        public void PageMargins()
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.PageMargins
            //ExSummary:Shows how to set the margins around pages in HTML file.
            Document doc = new Document(MyDir + "Bookmarks.docx");

            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions
            {
                PageMargins = 10
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.PageMargins.html", saveOptions);
            //ExEnd
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
            //ExFor:HtmlFixedSaveOptions.OptimizeOutput
            //ExSummary:Shows how to optimize document objects while saving to html.
            Document doc = new Document(MyDir + "Graphics.doc");

            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions { OptimizeOutput = false };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.OptimizeGraphicsOutput.html", saveOptions);
            //ExEnd
        }

        //ExStart
        //ExFor:ExportFontFormat
        //ExFor:HtmlFixedSaveOptions.FontFormat
        //ExFor:HtmlFixedSaveOptions.UseTargetMachineFonts
        //ExFor:IResourceSavingCallback
        //ExFor:IResourceSavingCallback.ResourceSaving(ResourceSavingArgs)
        //ExFor:ResourceSavingArgs
        //ExFor:ResourceSavingArgs.Document
        //ExFor:ResourceSavingArgs.KeepResourceStreamOpen
        //ExFor:ResourceSavingArgs.ResourceFileName
        //ExFor:ResourceSavingArgs.ResourceFileUri
        //ExFor:ResourceSavingArgs.ResourceStream
        //ExSummary:Shows how used target machine fonts to display the document.
        [Test] //ExSkip
        public void UsingMachineFonts()
        {
            Document doc = new Document(MyDir + "AltFontBulletPoints.docx");

            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions
            {
                UseTargetMachineFonts = true,
                FontFormat = ExportFontFormat.Ttf,
                ExportEmbeddedFonts = false,
                ResourceSavingCallback = new ResourceSavingCallback()
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.UsingMachineFonts.html", saveOptions);
        }

        private class ResourceSavingCallback : IResourceSavingCallback
        {
            /// <summary>
            /// Called when Aspose.Words saves an external resource to fixed page HTML or SVG.
            /// </summary>
            public void ResourceSaving(ResourceSavingArgs args)
            {
                Console.WriteLine($"Original document URI:\t{args.Document.OriginalFileName}");
                Console.WriteLine($"Resource being saved:\t{args.ResourceFileName}");
                Console.WriteLine($"Full uri after saving:\t{args.ResourceFileUri}");

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

        //ExStart
        //ExFor:HtmlFixedSaveOptions
        //ExFor:HtmlFixedSaveOptions.ResourceSavingCallback
        //ExFor:HtmlFixedSaveOptions.ResourcesFolder
        //ExFor:HtmlFixedSaveOptions.ResourcesFolderAlias
        //ExFor:HtmlFixedSaveOptions.SaveFormat
        //ExFor:HtmlFixedSaveOptions.ShowPageBorder
        //ExSummary:Shows how to print the URIs of linked resources created during conversion of a document to fixed-form .html.
        [Test] //ExSkip
        public void HtmlFixedResourceFolder()
        {
            // Open a document which contains images
            Document doc = new Document(MyDir + "Rendering.doc");

            HtmlFixedSaveOptions options = new HtmlFixedSaveOptions
            {
                SaveFormat = SaveFormat.HtmlFixed,
                ExportEmbeddedImages = false,
                ResourcesFolder = ArtifactsDir + "HtmlFixedResourceFolder",
                ResourcesFolderAlias = ArtifactsDir + "HtmlFixedResourceFolderAlias",
                ShowPageBorder = false,
                ResourceSavingCallback = new ResourceUriPrinter()
            };

            // A folder specified by ResourcesFolderAlias will contain the resources instead of ResourcesFolder
            // We must ensure the folder exists before the streams can put their resources into it
            Directory.CreateDirectory(options.ResourcesFolderAlias);

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.HtmlFixedResourceFolder.html", options);
        }

        /// <summary>
        /// Counts and prints URIs of resources contained by as they are converted to fixed .Html
        /// </summary>
        private class ResourceUriPrinter : IResourceSavingCallback
        {
            void IResourceSavingCallback.ResourceSaving(ResourceSavingArgs args)
            {
                // If we set a folder alias in the SaveOptions object, it will be printed here
                Console.WriteLine($"Resource #{++mSavedResourceCount} \"{args.ResourceFileName}\"");

                string extension = Path.GetExtension(args.ResourceFileName);
                switch (extension)
                {
                    case ".ttf":
                    case ".woff":
                    {
                        // By default 'ResourceFileUri' used system folder for fonts
                        // To avoid problems across platforms you must explicitly specify the path for the fonts
                        args.ResourceFileUri = ArtifactsDir + Path.DirectorySeparatorChar + args.ResourceFileName;
                        break;
                    }
                }
                Console.WriteLine("\t" + args.ResourceFileUri);

                // If we specified a ResourcesFolderAlias we will also need to redirect each stream to put its resource in that folder
                args.ResourceStream = new FileStream(args.ResourceFileUri, FileMode.Create);
                args.KeepResourceStreamOpen = false;
            }

            private int mSavedResourceCount;
        }
        //ExEnd
    }
}