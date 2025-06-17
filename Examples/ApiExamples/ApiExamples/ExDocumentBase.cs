// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Drawing;
using NUnit.Framework;
using Document = Aspose.Words.Document;
using IResourceLoadingCallback = Aspose.Words.Loading.IResourceLoadingCallback;
using SaveFormat = Aspose.Words.SaveFormat;
using System.IO;
using Aspose.Words.Loading;
using System.Net;
using Aspose.Pdf;
using System.Net.Http;

namespace ApiExamples
{
    [TestFixture]
    public class ExDocumentBase : ApiExampleBase
    {
        [Test]
        public void Constructor()
        {
            //ExStart
            //ExFor:DocumentBase
            //ExSummary:Shows how to initialize the subclasses of DocumentBase.
            Document doc = new Document();

            Assert.That(doc.GetType().BaseType, Is.EqualTo(typeof(DocumentBase)));

            GlossaryDocument glossaryDoc = new GlossaryDocument();
            doc.GlossaryDocument = glossaryDoc;

            Assert.That(glossaryDoc.GetType().BaseType, Is.EqualTo(typeof(DocumentBase)));
            //ExEnd
        }

        [Test]
        public void SetPageColor()
        {
            //ExStart
            //ExFor:DocumentBase.PageColor
            //ExSummary:Shows how to set the background color for all pages of a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            doc.PageColor = System.Drawing.Color.LightGray;

            doc.Save(ArtifactsDir + "DocumentBase.SetPageColor.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBase.SetPageColor.docx");

            Assert.That(doc.PageColor.ToArgb(), Is.EqualTo(System.Drawing.Color.LightGray.ToArgb()));
        }

        [Test]
        public void ImportNode()
        {
            //ExStart
            //ExFor:DocumentBase.ImportNode(Node, Boolean)
            //ExSummary:Shows how to import a node from one document to another.
            Document srcDoc = new Document();
            Document dstDoc = new Document();

            srcDoc.FirstSection.Body.FirstParagraph.AppendChild(
                new Run(srcDoc, "Source document first paragraph text."));
            dstDoc.FirstSection.Body.FirstParagraph.AppendChild(
                new Run(dstDoc, "Destination document first paragraph text."));

            // Every node has a parent document, which is the document that contains the node.
            // Inserting a node into a document that the node does not belong to will throw an exception.
            Assert.That(srcDoc.FirstSection.Document, Is.Not.EqualTo(dstDoc));
            Assert.Throws<ArgumentException>(() => dstDoc.AppendChild(srcDoc.FirstSection));

            // Use the ImportNode method to create a copy of a node, which will have the document
            // that called the ImportNode method set as its new owner document.
            Section importedSection = (Section)dstDoc.ImportNode(srcDoc.FirstSection, true);

            Assert.That(importedSection.Document, Is.EqualTo(dstDoc));

            // We can now insert the node into the document.
            dstDoc.AppendChild(importedSection);

            Assert.That(dstDoc.ToString(SaveFormat.Text), Is.EqualTo("Destination document first paragraph text.\r\nSource document first paragraph text.\r\n"));
            //ExEnd

            Assert.That(srcDoc.FirstSection, Is.Not.EqualTo(importedSection));
            Assert.That(srcDoc.FirstSection.Document, Is.Not.EqualTo(importedSection.Document));
            Assert.That(srcDoc.FirstSection.Body.FirstParagraph.GetText(), Is.EqualTo(importedSection.Body.FirstParagraph.GetText()));
        }

        [Test]
        public void ImportNodeCustom()
        {
            //ExStart
            //ExFor:DocumentBase.ImportNode(Node, Boolean, ImportFormatMode)
            //ExSummary:Shows how to import node from source document to destination document with specific options.
            // Create two documents and add a character style to each document.
            // Configure the styles to have the same name, but different text formatting.
            Document srcDoc = new Document();
            Style srcStyle = srcDoc.Styles.Add(StyleType.Character, "My style");
            srcStyle.Font.Name = "Courier New";
            DocumentBuilder srcBuilder = new DocumentBuilder(srcDoc);
            srcBuilder.Font.Style = srcStyle;
            srcBuilder.Writeln("Source document text.");

            Document dstDoc = new Document();
            Style dstStyle = dstDoc.Styles.Add(StyleType.Character, "My style");
            dstStyle.Font.Name = "Calibri";
            DocumentBuilder dstBuilder = new DocumentBuilder(dstDoc);
            dstBuilder.Font.Style = dstStyle;
            dstBuilder.Writeln("Destination document text.");

            // Import the Section from the destination document into the source document, causing a style name collision.
            // If we use destination styles, then the imported source text with the same style name
            // as destination text will adopt the destination style.
            Section importedSection = (Section)dstDoc.ImportNode(srcDoc.FirstSection, true, ImportFormatMode.UseDestinationStyles);
            Assert.That(importedSection.Body.Paragraphs[0].Runs[0].GetText().Trim(), Is.EqualTo("Source document text.")); //ExSkip
            Assert.That(dstDoc.Styles["My style_0"], Is.Null); //ExSkip
            Assert.That(importedSection.Body.FirstParagraph.Runs[0].Font.Name, Is.EqualTo(dstStyle.Font.Name));
            Assert.That(importedSection.Body.FirstParagraph.Runs[0].Font.StyleName, Is.EqualTo(dstStyle.Name));

            // If we use ImportFormatMode.KeepDifferentStyles, the source style is preserved,
            // and the naming clash resolves by adding a suffix.
            dstDoc.ImportNode(srcDoc.FirstSection, true, ImportFormatMode.KeepDifferentStyles);
            Assert.That(dstDoc.Styles["My style"].Font.Name, Is.EqualTo(dstStyle.Font.Name));
            Assert.That(dstDoc.Styles["My style_0"].Font.Name, Is.EqualTo(srcStyle.Font.Name));
            //ExEnd
        }

        [Test]
        public void BackgroundShape()
        {
            //ExStart
            //ExFor:DocumentBase.BackgroundShape
            //ExSummary:Shows how to set a background shape for every page of a document.
            Document doc = new Document();

            Assert.That(doc.BackgroundShape, Is.Null);

            // The only shape type that we can use as a background is a rectangle.
            Shape shapeRectangle = new Shape(doc, ShapeType.Rectangle);

            // There are two ways of using this shape as a page background.
            // 1 -  A flat color:
            shapeRectangle.FillColor = System.Drawing.Color.LightBlue;
            doc.BackgroundShape = shapeRectangle;

            doc.Save(ArtifactsDir + "DocumentBase.BackgroundShape.FlatColor.docx");

            // 2 -  An image:
            shapeRectangle = new Shape(doc, ShapeType.Rectangle);
            shapeRectangle.ImageData.SetImage(ImageDir + "Transparent background logo.png");

            // Adjust the image's appearance to make it more suitable as a watermark.
            shapeRectangle.ImageData.Contrast = 0.2;
            shapeRectangle.ImageData.Brightness = 0.7;

            doc.BackgroundShape = shapeRectangle;

            Assert.That(doc.BackgroundShape.HasImage, Is.True);

            Aspose.Words.Saving.PdfSaveOptions saveOptions = new Aspose.Words.Saving.PdfSaveOptions
            {
                CacheBackgroundGraphics = false
            };

            // Microsoft Word does not support shapes with images as backgrounds,
            // but we can still see these backgrounds in other save formats such as .pdf.
            doc.Save(ArtifactsDir + "DocumentBase.BackgroundShape.Image.pdf", saveOptions);
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBase.BackgroundShape.FlatColor.docx");

            Assert.That(doc.BackgroundShape.FillColor.ToArgb(), Is.EqualTo(System.Drawing.Color.LightBlue.ToArgb()));
            Assert.Throws<ArgumentException>(() =>
            {
                doc.BackgroundShape = new Shape(doc, ShapeType.Triangle);
            });
        }

        [Test]
        public void UsePdfDocumentForBackgroundShape()
        {
            BackgroundShape();

            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "DocumentBase.BackgroundShape.Image.pdf");
            XImage pdfDocImage = pdfDocument.Pages[1].Resources.Images[1];

            Assert.That(pdfDocImage.Width, Is.EqualTo(400));
            Assert.That(pdfDocImage.Height, Is.EqualTo(400));
            Assert.That(pdfDocImage.GetColorType(), Is.EqualTo(ColorType.Rgb));
        }

        //ExStart
        //ExFor:DocumentBase.ResourceLoadingCallback
        //ExFor:IResourceLoadingCallback
        //ExFor:IResourceLoadingCallback.ResourceLoading(ResourceLoadingArgs)
        //ExFor:ResourceLoadingAction
        //ExFor:ResourceLoadingArgs
        //ExFor:ResourceLoadingArgs.OriginalUri
        //ExFor:ResourceLoadingArgs.ResourceType
        //ExFor:ResourceLoadingArgs.SetData(Byte[])
        //ExFor:ResourceType
        //ExSummary:Shows how to customize the process of loading external resources into a document.
        [Test] //ExSkip
        public void ResourceLoadingCallback()
        {
            Document doc = new Document();
            doc.ResourceLoadingCallback = new ImageNameHandler();

            DocumentBuilder builder = new DocumentBuilder(doc);

            // Images usually are inserted using a URI, or a byte array.
            // Every instance of a resource load will call our callback's ResourceLoading method.
            builder.InsertImage("Google logo");
            builder.InsertImage("Aspose logo");
            builder.InsertImage("Watermark");

            Assert.That(doc.GetChildNodes(NodeType.Shape, true).Count, Is.EqualTo(3));

            doc.Save(ArtifactsDir + "DocumentBase.ResourceLoadingCallback.docx");
            TestResourceLoadingCallback(new Document(ArtifactsDir + "DocumentBase.ResourceLoadingCallback.docx")); //ExSkip
        }

        /// <summary>
        /// Allows us to load images into a document using predefined shorthands, as opposed to URIs.
        /// This will separate image loading logic from the rest of the document construction.
        /// </summary>
        private class ImageNameHandler : IResourceLoadingCallback
        {
            public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
            {
                // If this callback encounters one of the image shorthands while loading an image,
                // it will apply unique logic for each defined shorthand instead of treating it as a URI.
                if (args.ResourceType == ResourceType.Image)
                    switch (args.OriginalUri)
                    {
                        case "Google logo":
                            using (HttpClient client = new HttpClient())
                            {
                                byte[] imageData = client.GetByteArrayAsync("http://www.google.com/images/logos/ps_logo2.png").GetAwaiter().GetResult();
                                args.SetData(imageData);
                            }

                            return ResourceLoadingAction.UserProvided;

                        case "Aspose logo":
                            args.SetData(File.ReadAllBytes(ImageDir + "Logo.jpg"));

                            return ResourceLoadingAction.UserProvided;

                        case "Watermark":
                            args.SetData(File.ReadAllBytes(ImageDir + "Transparent background logo.png"));

                            return ResourceLoadingAction.UserProvided;
                    }

                return ResourceLoadingAction.Default;
            }
        }
        //ExEnd

        private void TestResourceLoadingCallback(Document doc)
        {
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                Assert.That(shape.HasImage, Is.True);
                Assert.That(shape.ImageData.ImageBytes, Is.Not.Empty);
            }
        }
    }
}
