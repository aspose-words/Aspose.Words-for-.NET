using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using NUnit.Framework;

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
            saveOptions.ImageSavingCallback = new SavedImageRename("DocumentBuilder.HandleDocument.md");

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
    }
}
