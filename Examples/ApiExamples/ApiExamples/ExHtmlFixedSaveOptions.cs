// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
            //ExSummary:Shows how to set which encoding to use while exporting a document to HTML.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Writeln("Hello World!");

            // The default encoding is UTF-8. If we want to represent our document using a different encoding,
            // we can use a SaveOptions object to set a specific encoding.
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

        [TestCase(true)]
        [TestCase(false)]
        public void ExportEmbeddedCss(bool exportEmbeddedCss)
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.ExportEmbeddedCss
            //ExSummary:Shows how to determine where to store CSS stylesheets when exporting a document to Html.
            Document doc = new Document(MyDir + "Rendering.docx");

            // When we export a document to html, Aspose.Words will also create a CSS stylesheet to format the document with.
            // Setting the "ExportEmbeddedCss" flag to "true" save the CSS stylesheet to a .css file,
            // and link to the file from the html document using a <link> element.
            // Setting the flag to "false" will embed the CSS stylesheet within the Html document,
            // which will create only one file instead of two.
            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                ExportEmbeddedCss = exportEmbeddedCss
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedCss.html", htmlFixedSaveOptions);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedCss.html");

            if (exportEmbeddedCss)
            {
                Assert.True(Regex.Match(outDocContents, "<style type=\"text/css\">").Success);
                Assert.False(File.Exists(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedCss/styles.css"));
            }
            else
            {
                Assert.True(Regex.Match(outDocContents,
                    "<link rel=\"stylesheet\" type=\"text/css\" href=\"HtmlFixedSaveOptions[.]ExportEmbeddedCss/styles[.]css\" media=\"all\" />").Success);
                Assert.True(File.Exists(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedCss/styles.css"));
            }
            //ExEnd
        }

        [TestCase(true)]
        [TestCase(false)]
        public void ExportEmbeddedFonts(bool exportEmbeddedFonts)
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.ExportEmbeddedFonts
            //ExSummary:Shows how to determine where to store embedded fonts when exporting a document to Html.
            Document doc = new Document(MyDir + "Embedded font.docx");

            // When we export a document with embedded fonts to .html,
            // Aspose.Words can place the fonts in two possible locations.
            // Setting the "ExportEmbeddedFonts" flag to "true" will store the raw data for embedded fonts within the CSS stylesheet,
            // in the "url" property of the "@font-face" rule. This may create a huge CSS stylesheet file
            // and reduce the number of external files that this HTML conversion will create.
            // Setting this flag to "false" will create a file for each font.
            // The CSS stylesheet will link to each font file using the "url" property of the "@font-face" rule.
            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                ExportEmbeddedFonts = exportEmbeddedFonts
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedFonts.html", htmlFixedSaveOptions);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedFonts/styles.css");

            if (exportEmbeddedFonts)
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

        [TestCase(true)]
        [TestCase(false)]
        public void ExportEmbeddedImages(bool exportImages)
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.ExportEmbeddedImages
            //ExSummary:Shows how to determine where to store images when exporting a document to Html.
            Document doc = new Document(MyDir + "Images.docx");

            // When we export a document with embedded images to .html,
            // Aspose.Words can place the images in two possible locations.
            // Setting the "ExportEmbeddedImages" flag to "true" will store the raw data
            // for all images within the output HTML document, in the "src" attribute of <image> tags.
            // Setting this flag to "false" will create an image file in the local file system for every image,
            // and store all these files in a separate folder.
            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                ExportEmbeddedImages = exportImages
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedImages.html", htmlFixedSaveOptions);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedImages.html");

            if (exportImages)
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

        [TestCase(true)]
        [TestCase(false)]
        public void ExportEmbeddedSvgs(bool exportSvgs)
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.ExportEmbeddedSvg
            //ExSummary:Shows how to determine where to store SVG objects when exporting a document to Html.
            Document doc = new Document(MyDir + "Images.docx");

            // When we export a document with SVG objects to .html,
            // Aspose.Words can place these objects in two possible locations.
            // Setting the "ExportEmbeddedSvg" flag to "true" will embed all SVG object raw data
            // within the output HTML, inside <image> tags.
            // Setting this flag to "false" will create a file in the local file system for each SVG object.
            // The HTML will link to each file using the "data" attribute of an <object> tag.
            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                ExportEmbeddedSvg = exportSvgs
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedSvgs.html", htmlFixedSaveOptions);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlFixedSaveOptions.ExportEmbeddedSvgs.html");

            if (exportSvgs)
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

        [TestCase(true)]
        [TestCase(false)]
        public void ExportFormFields(bool exportFormFields)
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.ExportFormFields
            //ExSummary:Shows how to export form fields to Html.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertCheckBox("CheckBox", false, 15);

            // When we export a document with form fields to .html,
            // there are two ways in which Aspose.Words can export form fields.
            // Setting the "ExportFormFields" flag to "true" will export them as interactive objects.
            // Setting this flag to "false" will display form fields as plain text.
            // This will freeze them at their current value, and prevent the reader of our HTML document
            // from being able to interact with them.
            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                ExportFormFields = exportFormFields
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.ExportFormFields.html", htmlFixedSaveOptions);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlFixedSaveOptions.ExportFormFields.html");

            if (exportFormFields)
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
            //ExSummary:Shows how to place CSS into a separate file and add a prefix to all of its CSS class names.
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

            outDocContents = File.ReadAllText(ArtifactsDir + "HtmlFixedSaveOptions.AddCssClassNamesPrefix/styles.css");

            Assert.True(Regex.Match(outDocContents,
                ".myprefixdiv { position:absolute; } " +
                ".myprefixspan { position:absolute; white-space:pre; color:#000000; font-size:12pt; }").Success);
            //ExEnd
        }

        [TestCase(HtmlFixedPageHorizontalAlignment.Center)]
        [TestCase(HtmlFixedPageHorizontalAlignment.Left)]
        [TestCase(HtmlFixedPageHorizontalAlignment.Right)]
        public void HorizontalAlignment(HtmlFixedPageHorizontalAlignment pageHorizontalAlignment)
        {
            //ExStart
            //ExFor:HtmlFixedSaveOptions.PageHorizontalAlignment
            //ExFor:HtmlFixedPageHorizontalAlignment
            //ExSummary:Shows how to set the horizontal alignment of pages when saving a document to HTML.
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
            //ExSummary:Shows how to adjust page margins when saving a document to HTML.
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

        [TestCase(false)]
        [TestCase(true)]
        public void OptimizeGraphicsOutput(bool optimizeOutput)
        {
            //ExStart
            //ExFor:FixedPageSaveOptions.OptimizeOutput
            //ExFor:HtmlFixedSaveOptions.OptimizeOutput
            //ExSummary:Shows how to simplify a document when saving it to HTML by removing various redundant objects.
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions { OptimizeOutput = optimizeOutput };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.OptimizeGraphicsOutput.html", saveOptions);

            // The size of the optimized version of the document is almost a third of the size of the unoptimized document.
            if (optimizeOutput)
                Assert.AreEqual(58000, 
                    new FileInfo(ArtifactsDir + "HtmlFixedSaveOptions.OptimizeGraphicsOutput.html").Length, 200);
            else
                Assert.AreEqual(161100, 
                    new FileInfo(ArtifactsDir + "HtmlFixedSaveOptions.OptimizeGraphicsOutput.html").Length, 200);
            //ExEnd
        }


        [TestCase(false)]
        [TestCase(true)]
        public void UsingMachineFonts(bool useTargetMachineFonts)
        {
            //ExStart
            //ExFor:ExportFontFormat
            //ExFor:HtmlFixedSaveOptions.FontFormat
            //ExFor:HtmlFixedSaveOptions.UseTargetMachineFonts
            //ExSummary:Shows how use fonts only from the target machine when saving a document to HTML.
            Document doc = new Document(MyDir + "Bullet points with alternative font.docx");

            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions
            {
                ExportEmbeddedCss = true,
                UseTargetMachineFonts = useTargetMachineFonts,
                FontFormat = ExportFontFormat.Ttf,
                ExportEmbeddedFonts = false,
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.UsingMachineFonts.html", saveOptions);

            string outDocContents = File.ReadAllText(ArtifactsDir + "HtmlFixedSaveOptions.UsingMachineFonts.html");

            if (useTargetMachineFonts)
                Assert.False(Regex.Match(outDocContents, "@font-face").Success);
            else
                Assert.True(Regex.Match(outDocContents,
                    "@font-face { font-family:'Arial'; font-style:normal; font-weight:normal; src:local[(]'☺'[)], " +
                    "url[(]'HtmlFixedSaveOptions.UsingMachineFonts/font001.ttf'[)] format[(]'truetype'[)]; }").Success);
            //ExEnd
        }

        //ExStart
        //ExFor:IResourceSavingCallback
        //ExFor:IResourceSavingCallback.ResourceSaving(ResourceSavingArgs)
        //ExFor:ResourceSavingArgs
        //ExFor:ResourceSavingArgs.Document
        //ExFor:ResourceSavingArgs.ResourceFileName
        //ExFor:ResourceSavingArgs.ResourceFileUri
        //ExSummary:Shows how to use a callback to track external resources created while converting a document to HTML.
        [Test] //ExSkip
        public void ResourceSavingCallback()
        {
            Document doc = new Document(MyDir + "Bullet points with alternative font.docx");

            FontSavingCallback callback = new FontSavingCallback();

            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions
            {
                ResourceSavingCallback = callback
            };

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.UsingMachineFonts.html", saveOptions);

            Console.WriteLine(callback.GetText());
            TestResourceSavingCallback(callback); //ExSkip
        }

        private class FontSavingCallback : IResourceSavingCallback
        {
            /// <summary>
            /// Called when Aspose.Words saves an external resource to fixed page HTML or SVG.
            /// </summary>
            public void ResourceSaving(ResourceSavingArgs args)
            {
                mText.AppendLine($"Original document URI:\t{args.Document.OriginalFileName}");
                mText.AppendLine($"Resource being saved:\t{args.ResourceFileName}");
                mText.AppendLine($"Full uri after saving:\t{args.ResourceFileUri}\n");
            }

            public string GetText()
            {
                return mText.ToString();
            }

            private readonly StringBuilder mText = new StringBuilder();
        }
        //ExEnd

        private void TestResourceSavingCallback(FontSavingCallback callback)
        {
            Assert.True(callback.GetText().Contains("font001.woff")); 
            Assert.True(callback.GetText().Contains("styles.css"));
        }

        //ExStart
        //ExFor:HtmlFixedSaveOptions
        //ExFor:HtmlFixedSaveOptions.ResourceSavingCallback
        //ExFor:HtmlFixedSaveOptions.ResourcesFolder
        //ExFor:HtmlFixedSaveOptions.ResourcesFolderAlias
        //ExFor:HtmlFixedSaveOptions.SaveFormat
        //ExFor:HtmlFixedSaveOptions.ShowPageBorder
        //ExFor:IResourceSavingCallback
        //ExFor:IResourceSavingCallback.ResourceSaving(ResourceSavingArgs)
        //ExFor:ResourceSavingArgs.KeepResourceStreamOpen
        //ExFor:ResourceSavingArgs.ResourceStream
        //ExSummary:Shows how to use a callback to print the URIs of external resources created while converting a document to HTML.
        [Test] //ExSkip
        public void HtmlFixedResourceFolder()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            ResourceUriPrinter callback = new ResourceUriPrinter();

            HtmlFixedSaveOptions options = new HtmlFixedSaveOptions
            {
                SaveFormat = SaveFormat.HtmlFixed,
                ExportEmbeddedImages = false,
                ResourcesFolder = ArtifactsDir + "HtmlFixedResourceFolder",
                ResourcesFolderAlias = ArtifactsDir + "HtmlFixedResourceFolderAlias",
                ShowPageBorder = false,
                ResourceSavingCallback = callback
            };

            // A folder specified by ResourcesFolderAlias will contain the resources instead of ResourcesFolder.
            // We must ensure the folder exists before the streams can put their resources into it.
            Directory.CreateDirectory(options.ResourcesFolderAlias);

            doc.Save(ArtifactsDir + "HtmlFixedSaveOptions.HtmlFixedResourceFolder.html", options);

            Console.WriteLine(callback.GetText());

            string[] resourceFiles = Directory.GetFiles(ArtifactsDir + "HtmlFixedResourceFolderAlias");

            Assert.False(Directory.Exists(ArtifactsDir + "HtmlFixedResourceFolder"));
            Assert.AreEqual(6, resourceFiles.Count(f => f.EndsWith(".jpeg") || f.EndsWith(".png") || f.EndsWith(".css")));
            TestHtmlFixedResourceFolder(callback); //ExSkip
        }
        
        /// <summary>
        /// Counts and prints URIs of resources contained by as they are converted to fixed HTML.
        /// </summary>
        private class ResourceUriPrinter : IResourceSavingCallback
        {
            void IResourceSavingCallback.ResourceSaving(ResourceSavingArgs args)
            {
                // If we set a folder alias in the SaveOptions object, we will be able to print it from here.
                mText.AppendLine($"Resource #{++mSavedResourceCount} \"{args.ResourceFileName}\"");

                string extension = Path.GetExtension(args.ResourceFileName);
                switch (extension)
                {
                    case ".ttf":
                    case ".woff":
                    {
                        // By default, 'ResourceFileUri' uses system folder for fonts.
                        // To avoid problems in other platforms you must explicitly specify the path for the fonts.
                        args.ResourceFileUri = ArtifactsDir + Path.DirectorySeparatorChar + args.ResourceFileName;
                        break;
                    }
                }

                mText.AppendLine("\t" + args.ResourceFileUri);

                // If we have specified a folder in the "ResourcesFolderAlias" property,
                // we will also need to redirect each stream to put its resource in that folder.
                args.ResourceStream = new FileStream(args.ResourceFileUri, FileMode.Create);
                args.KeepResourceStreamOpen = false;
            }

            public string GetText()
            {
                return mText.ToString();
            }

            private int mSavedResourceCount;
            private readonly StringBuilder mText = new StringBuilder();
        }
        //ExEnd

        private void TestHtmlFixedResourceFolder(ResourceUriPrinter callback)
        {
            Assert.AreEqual(16, Regex.Matches(callback.GetText(), "Resource #").Count);
        }
    }
}