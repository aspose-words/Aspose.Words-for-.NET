// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Text;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
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

            // The default encoding is UTF-8
            // If we want to represent our document using a different encoding, we can set one explicitly using a SaveOptions object
            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                Encoding = Encoding.GetEncoding("ASCII")
            };

            Assert.AreEqual("US-ASCII", htmlFixedSaveOptions.Encoding.EncodingName);

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.UseEncoding.html", htmlFixedSaveOptions);
            //ExEnd

            Assert.True(Regex.Match(File.ReadAllText(ArtifactsDir + "HtmlFixedSaveOptions.UseEncoding.html"), 
                "content=\"text/html; charset=us-ascii\"").Success);
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

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void ExportEmbeddedCSS(bool doExportEmbeddedCss)
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.ExportEmbeddedCss
            //ExSummary:Shows how to export embedded stylesheets into an HTML file.
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                ExportEmbeddedCss = doExportEmbeddedCss
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedCSS.html", htmlFixedSaveOptions);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedCSS.html");

            if (doExportEmbeddedCss)
            {
                Assert.True(Regex.Match(outDocContents, "<style type=\"text/css\">").Success);
                Assert.False(File.Exists(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedCSS/styles.css"));
            }
            else
            {
                Assert.True(Regex.Match(outDocContents,
                    "<link rel=\"stylesheet\" type=\"text/css\" href=\"HtmlFixedSaveOptions[.]ExportEmbeddedCSS/styles[.]css\" media=\"all\" />").Success);
                Assert.True(File.Exists(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedCSS/styles.css"));
            }
            //ExEnd
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void ExportEmbeddedFonts(bool doExportEmbeddedFonts)
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.ExportEmbeddedFonts
            //ExSummary:Shows how to export embedded fonts into an HTML file.
            Document doc = new Document(MyDir + "Embedded font.docx");

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                ExportEmbeddedFonts = doExportEmbeddedFonts
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedFonts.html", htmlFixedSaveOptions);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedFonts/styles.css");

            if (doExportEmbeddedFonts)
            {
                Assert.True(Regex.Match(outDocContents,
                    "@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], url[(].+[)] format[(]'woff'[)]; }").Success);
                Assert.AreEqual(0, Directory.GetFiles(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedFonts").Count(f => f.EndsWith(".woff")));
            }
            else
            {
                Assert.True(Regex.Match(outDocContents,
                    "@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], url[(]'font001[.]woff'[)] format[(]'woff'[)]; }").Success);
                Assert.AreEqual(2, Directory.GetFiles(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedFonts").Count(f => f.EndsWith(".woff")));
            }
            //ExEnd
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void ExportEmbeddedImages(bool doExportImages)
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.ExportEmbeddedImages
            //ExSummary:Shows how to export embedded images into an HTML file.
            Document doc = new Document(MyDir + "Images.docx");

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                ExportEmbeddedImages = doExportImages
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedImages.html", htmlFixedSaveOptions);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedImages.html");

            if (doExportImages)
            {
                Assert.False(File.Exists(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedImages/image001.jpeg"));
                Assert.True(Regex.Match(outDocContents,
                    "<img class=\"awimg\" style=\"left:0pt; top:0pt; width:493.1pt; height:300.55pt;\" src=\".+\" />").Success);
            }
            else
            {
                Assert.True(File.Exists(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedImages/image001.jpeg"));
                Assert.True(Regex.Match(outDocContents,
                    "<img class=\"awimg\" style=\"left:0pt; top:0pt; width:493.1pt; height:300.55pt;\" " +
                    "src=\"HtmlFixedSaveOptions[.]ExportEmbeddedImages/image001[.]jpeg\" />").Success);
            }
            //ExEnd
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void ExportEmbeddedSvgs(bool doExportSvgs)
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.ExportEmbeddedSvg
            //ExSummary:Shows how to export embedded SVG objects into an HTML file.
            Document doc = new Document(MyDir + "Images.docx");

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                ExportEmbeddedSvg = doExportSvgs
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedSvgs.html", htmlFixedSaveOptions);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedSvgs.html");

            if (doExportSvgs)
            {
                Assert.False(File.Exists(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedSvgs/svg001.svg"));
                Assert.True(Regex.Match(outDocContents,
                    "<image id=\"image004\" xlink:href=.+/>").Success);
            }
            else
            {
                Assert.True(File.Exists(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedSvgs/svg001.svg"));
                Assert.True(Regex.Match(outDocContents,
                    "<object type=\"image/svg[+]xml\" data=\"HtmlFixedSaveOptions.ExportEmbeddedSvgs/svg001[.]svg\"></object>").Success);
            }
            //ExEnd
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void ExportFormFields(bool doExportFormFields)
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.ExportFormFields
            //ExSummary:Show how to exporting form fields from a document into HTML file.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertCheckBox("CheckBox", false, 15);

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                ExportFormFields = doExportFormFields
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.ExportFormFields.html", htmlFixedSaveOptions);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlFixedSaveOptions.ExportFormFields.html");

            if (doExportFormFields)
            {
                Assert.True(Regex.Match(outDocContents,
                    "<a name=\"CheckBox\" style=\"left:0pt; top:0pt;\"></a>" +
                    "<input style=\"position:absolute; left:0pt; top:0pt;\" type=\"checkbox\" name=\"CheckBox\" />").Success);
            }
            else
            {
                Assert.True(Regex.Match(outDocContents, 
                    "<a name=\"CheckBox\" style=\"left:0pt; top:0pt;\"></a>" +
                    "<div class=\"awdiv\" style=\"left:0.8pt; top:0.8pt; width:14.25pt; height:14.25pt; border:solid 0.75pt #000000;\"").Success);
            }
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
                CssClassNamesPrefix = "myprefix",
                SaveFontFaceCssSeparately = true
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.AddCssClassNamesPrefix.html", htmlFixedSaveOptions);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlFixedSaveOptions.AddCssClassNamesPrefix.html");

            Assert.True(Regex.Match(outDocContents,
                "<div class=\"myprefixdiv myprefixpage\" style=\"width:595[.]3pt; height:841[.]9pt;\">" +
                "<div class=\"myprefixdiv\" style=\"left:85[.]05pt; top:36pt; clip:rect[(]0pt,510[.]25pt,74[.]95pt,-85.05pt[)];\">" +
                "<span class=\"myprefixspan myprefixtext001\" style=\"font-size:11pt; left:294[.]73pt; top:0[.]36pt;\">").Success);
            //ExEnd
        }

        [Test]
        [TestCase(HtmlFixedPageHorizontalAlignment.Center)]
        [TestCase(HtmlFixedPageHorizontalAlignment.Left)]
        [TestCase(HtmlFixedPageHorizontalAlignment.Right)]
        public void HorizontalAlignment(HtmlFixedPageHorizontalAlignment pageHorizontalAlignment)
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.PageHorizontalAlignment
            //ExFor:HtmlFixedPageHorizontalAlignment
            //ExSummary:Shows how to set the horizontal alignment of pages in HTML file.
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                PageHorizontalAlignment = pageHorizontalAlignment
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.HorizontalAlignment.html", htmlFixedSaveOptions);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlFixedSaveOptions.HorizontalAlignment/styles.css");

            switch (pageHorizontalAlignment)
            {
                case HtmlFixedPageHorizontalAlignment.Center:
                    Assert.True(Regex.Match(outDocContents,
                        "[.]awpage { position:relative; border:solid 1pt black; margin:10pt auto 10pt auto; overflow:hidden; }").Success);
                    break;
                case HtmlFixedPageHorizontalAlignment.Left:
                    Assert.True(Regex.Match(outDocContents, 
                        "[.]awpage { position:relative; border:solid 1pt black; margin:10pt auto 10pt 10pt; overflow:hidden; }").Success);
                    break;
                case HtmlFixedPageHorizontalAlignment.Right:
                    Assert.True(Regex.Match(outDocContents, 
                        "[.]awpage { position:relative; border:solid 1pt black; margin:10pt 10pt 10pt auto; overflow:hidden; }").Success);
                    break;
            }
            //ExEnd
        }

        [Test]
        public void PageMargins()
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.PageMargins
            //ExSummary:Shows how to set the margins around pages in HTML file.
            Document doc = new Document(MyDir + "Document.docx");

            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions
            {
                PageMargins = 15
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.PageMargins.html", saveOptions);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlFixedSaveOptions.PageMargins/styles.css");

            Assert.True(Regex.Match(outDocContents,
                "[.]awpage { position:relative; border:solid 1pt black; margin:15pt auto 15pt auto; overflow:hidden; }").Success);
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
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions { OptimizeOutput = false };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.OptimizeGraphicsOutput.Unoptimized.html", saveOptions);

            saveOptions.OptimizeOutput = true;

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.OptimizeGraphicsOutput.Optimized.html", saveOptions);

            Assert.True(new FileInfo(ArtifactsDir + "HtmlFixedSaveOptions.OptimizeGraphicsOutput.Unoptimized.html").Length > 
                            new FileInfo(ArtifactsDir + "HtmlFixedSaveOptions.OptimizeGraphicsOutput.Optimized.html").Length);
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
        //ExSummary:Shows how use target machine fonts to display the document.
        [Test] //ExSkip
        public void UsingMachineFonts()
        {
            Document doc = new Document(MyDir + "Bullet points with alternative font.docx");

            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions
            {
                ExportEmbeddedCss = true,
                UseTargetMachineFonts = true,
                FontFormat = ExportFontFormat.Ttf,
                ExportEmbeddedFonts = false,
                ResourceSavingCallback = new ResourceSavingCallback()
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.UsingMachineFonts.html", saveOptions);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlFixedSaveOptions.UsingMachineFonts.html");

            if (saveOptions.UseTargetMachineFonts)
                Assert.False(Regex.Match(outDocContents, "@font-face").Success);
            else
                Assert.True(Regex.Match(outDocContents,
                    "@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], " +
                    "url[(]'HtmlFixedSaveOptions.UsingMachineFonts/font001.ttf'[)] format[(]'truetype'[)]; }").Success);
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
                        Assert.Fail("'ResourceSavingCallback' is not fired for fonts when 'UseTargetMachineFonts' is true");
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
            Document doc = new Document(MyDir + "Rendering.docx");

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

            string[] resourceFiles = Directory.GetFiles(ArtifactsDir + "HtmlFixedResourceFolderAlias");

            Assert.False(Directory.Exists(ArtifactsDir + "HtmlFixedResourceFolder"));
            Assert.AreEqual(6, resourceFiles.Count(f => f.EndsWith(".jpeg") || f.EndsWith(".png") || f.EndsWith(".css")));
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