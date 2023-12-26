// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
using System.Drawing;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using NUnit.Framework;
#if NET5_0_OR_GREATER || __MOBILE__
using SkiaSharp;
#endif

namespace ApiExamples
{
    class ExMarkdownSaveOptions : ApiExampleBase
    {
        [TestCase(TableContentAlignment.Left)]
        [TestCase(TableContentAlignment.Right)]
        [TestCase(TableContentAlignment.Center)]
        [TestCase(TableContentAlignment.Auto)]
        public void MarkdownDocumentTableContentAlignment(TableContentAlignment tableContentAlignment)
        {
            DocumentBuilder builder = new DocumentBuilder();

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            builder.Write("Cell1");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write("Cell2");

            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions { TableContentAlignment = tableContentAlignment };

            builder.Document.Save(ArtifactsDir + "MarkdownSaveOptions.MarkdownDocumentTableContentAlignment.md", saveOptions);

            Document doc = new Document(ArtifactsDir + "MarkdownSaveOptions.MarkdownDocumentTableContentAlignment.md");
            Table table = doc.FirstSection.Body.Tables[0];

            switch (tableContentAlignment)
            {
                case TableContentAlignment.Auto:
                    Assert.AreEqual(ParagraphAlignment.Right,
                        table.FirstRow.Cells[0].FirstParagraph.ParagraphFormat.Alignment);
                    Assert.AreEqual(ParagraphAlignment.Center,
                        table.FirstRow.Cells[1].FirstParagraph.ParagraphFormat.Alignment);
                    break;
                case TableContentAlignment.Left:
                    Assert.AreEqual(ParagraphAlignment.Left,
                        table.FirstRow.Cells[0].FirstParagraph.ParagraphFormat.Alignment);
                    Assert.AreEqual(ParagraphAlignment.Left,
                        table.FirstRow.Cells[1].FirstParagraph.ParagraphFormat.Alignment);
                    break;
                case TableContentAlignment.Center:
                    Assert.AreEqual(ParagraphAlignment.Center,
                        table.FirstRow.Cells[0].FirstParagraph.ParagraphFormat.Alignment);
                    Assert.AreEqual(ParagraphAlignment.Center,
                        table.FirstRow.Cells[1].FirstParagraph.ParagraphFormat.Alignment);
                    break;
                case TableContentAlignment.Right:
                    Assert.AreEqual(ParagraphAlignment.Right,
                        table.FirstRow.Cells[0].FirstParagraph.ParagraphFormat.Alignment);
                    Assert.AreEqual(ParagraphAlignment.Right,
                        table.FirstRow.Cells[1].FirstParagraph.ParagraphFormat.Alignment);
                    break;
            }
        }

        //ExStart
        //ExFor:MarkdownSaveOptions.ImageSavingCallback
        //ExFor:IImageSavingCallback
        //ExSummary:Shows how to rename the image name during saving into Markdown document.
        [Test] //ExSkip
        public void RenameImages()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();

            // If we convert a document that contains images into Markdown, we will end up with one Markdown file which links to several images.
            // Each image will be in the form of a file in the local file system.
            // There is also a callback that can customize the name and file system location of each image.
            saveOptions.ImageSavingCallback = new SavedImageRename("MarkdownSaveOptions.HandleDocument.md");

            // The ImageSaving() method of our callback will be run at this time.
            doc.Save(ArtifactsDir + "MarkdownSaveOptions.HandleDocument.md", saveOptions);

            Assert.AreEqual(1,
                Directory.GetFiles(ArtifactsDir)
                    .Where(s => s.StartsWith(ArtifactsDir + "MarkdownSaveOptions.HandleDocument.md shape"))
                    .Count(f => f.EndsWith(".jpeg")));
            Assert.AreEqual(8,
                Directory.GetFiles(ArtifactsDir)
                    .Where(s => s.StartsWith(ArtifactsDir + "MarkdownSaveOptions.HandleDocument.md shape"))
                    .Count(f => f.EndsWith(".png")));
        }

        /// <summary>
        /// Renames saved images that are produced when an Markdown document is saved.
        /// </summary>
        public class SavedImageRename : IImageSavingCallback
        {
            public SavedImageRename(string outFileName)
            {
                mOutFileName = outFileName;
            }

            void IImageSavingCallback.ImageSaving(ImageSavingArgs args)
            {
                string imageFileName = $"{mOutFileName} shape {++mCount}, of type {args.CurrentShape.ShapeType}{Path.GetExtension(args.ImageFileName)}";

                args.ImageFileName = imageFileName;
                args.ImageStream = new FileStream(ArtifactsDir + imageFileName, FileMode.Create);

                Assert.True(args.ImageStream.CanWrite);
                Assert.True(args.IsImageAvailable);
                Assert.False(args.KeepImageStreamOpen);
            }

            private int mCount;
            private readonly string mOutFileName;
        }
        //ExEnd

        [TestCase(true)]
        [TestCase(false)]
        public void ExportImagesAsBase64(bool exportImagesAsBase64)
        {
            //ExStart
            //ExFor:MarkdownSaveOptions.ExportImagesAsBase64
            //ExSummary:Shows how to save a .md document with images embedded inside it.
            Document doc = new Document(MyDir + "Images.docx");

            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions { ExportImagesAsBase64 = exportImagesAsBase64 };

            doc.Save(ArtifactsDir + "MarkdownSaveOptions.ExportImagesAsBase64.md", saveOptions);

            string outDocContents = File.ReadAllText(ArtifactsDir + "MarkdownSaveOptions.ExportImagesAsBase64.md");

            Assert.True(exportImagesAsBase64
                ? outDocContents.Contains("data:image/jpeg;base64")
                : outDocContents.Contains("MarkdownSaveOptions.ExportImagesAsBase64.001.jpeg"));
            //ExEnd
        }

        [TestCase(MarkdownListExportMode.PlainText)]
        [TestCase(MarkdownListExportMode.MarkdownSyntax)]
        public void ListExportMode(MarkdownListExportMode markdownListExportMode)
        {
            //ExStart
            //ExFor:MarkdownSaveOptions.ListExportMode
            //ExSummary:Shows how to list items will be written to the markdown document.
            Document doc = new Document(MyDir + "List item.docx");

            // Use MarkdownListExportMode.PlainText or MarkdownListExportMode.MarkdownSyntax to export list.
            MarkdownSaveOptions options = new MarkdownSaveOptions { ListExportMode = markdownListExportMode };
            doc.Save(ArtifactsDir + "MarkdownSaveOptions.ListExportMode.md", options);
            //ExEnd
        }

        [Test]
        public void ImagesFolder()
        {
            //ExStart
            //ExFor:MarkdownSaveOptions.ImagesFolder
            //ExFor:MarkdownSaveOptions.ImagesFolderAlias
            //ExSummary:Shows how to specifies the name of the folder used to construct image URIs.
            DocumentBuilder builder = new DocumentBuilder();

            builder.Writeln("Some image below:");            
            builder.InsertImage(ImageDir + "Logo.jpg");

            string imagesFolder = Path.Combine(ArtifactsDir, "ImagesDir");
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
            // Use the "ImagesFolder" property to assign a folder in the local file system into which
            // Aspose.Words will save all the document's linked images.
            saveOptions.ImagesFolder = imagesFolder;
            // Use the "ImagesFolderAlias" property to use this folder
            // when constructing image URIs instead of the images folder's name.
            saveOptions.ImagesFolderAlias = "http://example.com/images";

            builder.Document.Save(ArtifactsDir + "MarkdownSaveOptions.ImagesFolder.md", saveOptions);
            //ExEnd

            string[] dirFiles = Directory.GetFiles(imagesFolder, "MarkdownSaveOptions.ImagesFolder.001.jpeg");
            Assert.AreEqual(1, dirFiles.Length);
            Document doc = new Document(ArtifactsDir + "MarkdownSaveOptions.ImagesFolder.md");
            doc.GetText().Contains("http://example.com/images/MarkdownSaveOptions.ImagesFolder.001.jpeg");
        }

        [Test]
        public void ExportUnderlineFormatting()
        {
            //ExStart:ExportUnderlineFormatting
            //ReleaseVersion:24.01
            //ExFor:MarkdownSaveOptions.ExportUnderlineFormatting
            //ExSummary:Shows how to export underline formatting as ++.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Underline = Underline.Single;
            builder.Write("Lorem ipsum. Dolor sit amet.");

            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions() { ExportUnderlineFormatting = true };
            doc.Save(ArtifactsDir + "MarkdownSaveOptions.ExportUnderlineFormatting.md", saveOptions);
            //ExEnd:ExportUnderlineFormatting
        }
    }
}
