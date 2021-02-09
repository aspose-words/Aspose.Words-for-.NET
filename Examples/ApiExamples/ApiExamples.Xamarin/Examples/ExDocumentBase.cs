// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
#if NET462 || JAVA
using System.IO;
using Aspose.Words.Loading;
using System.Net;
#endif
#if NET462 || NETCOREAPP2_1 || JAVA
using Aspose.Pdf;
#endif

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

            Assert.AreEqual(typeof(DocumentBase), doc.GetType().BaseType);

            GlossaryDocument glossaryDoc = new GlossaryDocument();
            doc.GlossaryDocument = glossaryDoc;

            Assert.AreEqual(typeof(DocumentBase), glossaryDoc.GetType().BaseType);
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

            Assert.AreEqual(System.Drawing.Color.LightGray.ToArgb(), doc.PageColor.ToArgb());
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
            Assert.AreNotEqual(dstDoc, srcDoc.FirstSection.Document);
            Assert.Throws<ArgumentException>(() => { dstDoc.AppendChild(srcDoc.FirstSection); });

            // Use the ImportNode method to create a copy of a node, which will have the document
            // that called the ImportNode method set as its new owner document.
            Section importedSection = (Section)dstDoc.ImportNode(srcDoc.FirstSection, true);

            Assert.AreEqual(dstDoc, importedSection.Document);

            // We can now insert the node into the document.
            dstDoc.AppendChild(importedSection);

            Assert.AreEqual("Destination document first paragraph text.\r\nSource document first paragraph text.\r\n",
                dstDoc.ToString(SaveFormat.Text));
            //ExEnd

            Assert.AreNotEqual(importedSection, srcDoc.FirstSection);
            Assert.AreNotEqual(importedSection.Document, srcDoc.FirstSection.Document);
            Assert.AreEqual(importedSection.Body.FirstParagraph.GetText(),
                srcDoc.FirstSection.Body.FirstParagraph.GetText());
        }

        [Test]
        public void ImportNodeCustom()
        {
            //ExStart
            //ExFor:DocumentBase.ImportNode(Node, System.Boolean, ImportFormatMode)
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
            Assert.AreEqual("Source document text.", importedSection.Body.Paragraphs[0].Runs[0].GetText().Trim()); //ExSkip
            Assert.IsNull(dstDoc.Styles["My style_0"]); //ExSkip
            Assert.AreEqual(dstStyle.Font.Name, importedSection.Body.FirstParagraph.Runs[0].Font.Name);
            Assert.AreEqual(dstStyle.Name, importedSection.Body.FirstParagraph.Runs[0].Font.StyleName);

            // If we use ImportFormatMode.KeepDifferentStyles, the source style is preserved,
            // and the naming clash resolves by adding a suffix.
            dstDoc.ImportNode(srcDoc.FirstSection, true, ImportFormatMode.KeepDifferentStyles);
            Assert.AreEqual(dstStyle.Font.Name, dstDoc.Styles["My style"].Font.Name);
            Assert.AreEqual(srcStyle.Font.Name, dstDoc.Styles["My style_0"].Font.Name);
            //ExEnd
        }

        [Test]
        public void BackgroundShape()
        {
            //ExStart
            //ExFor:DocumentBase.BackgroundShape
            //ExSummary:Shows how to set a background shape for every page of a document.
            Document doc = new Document();

            Assert.IsNull(doc.BackgroundShape);

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

            Assert.IsTrue(doc.BackgroundShape.HasImage);

            // Microsoft Word does not support shapes with images as backgrounds,
            // but we can still see these backgrounds in other save formats such as .pdf.
            doc.Save(ArtifactsDir + "DocumentBase.BackgroundShape.Image.pdf");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBase.BackgroundShape.FlatColor.docx");

            Assert.AreEqual(System.Drawing.Color.LightBlue.ToArgb(), doc.BackgroundShape.FillColor.ToArgb());
            Assert.Throws<ArgumentException>(() =>
            {
                doc.BackgroundShape = new Shape(doc, ShapeType.Triangle);
            });

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "DocumentBase.BackgroundShape.Image.pdf");
            XImage pdfDocImage = pdfDocument.Pages[1].Resources.Images[1];

            Assert.AreEqual(400, pdfDocImage.Width);
            Assert.AreEqual(400, pdfDocImage.Height);
            Assert.AreEqual(ColorType.Rgb, pdfDocImage.GetColorType());
#endif
        }

#if NET462 || JAVA
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

            Assert.AreEqual(3, doc.GetChildNodes(NodeType.Shape, true).Count);

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
                            using (WebClient webClient = new WebClient())
                            {
                                args.SetData(webClient.DownloadData("http://www.google.com/images/logos/ps_logo2.png"));
                            }

                            return ResourceLoadingAction.UserProvided;

                        case "Aspose logo":
                            using (WebClient webClient = new WebClient())
                            {
                                args.SetData(webClient.DownloadData(AsposeLogoUrl));
                            }

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
                Assert.IsTrue(shape.HasImage);
                Assert.IsNotEmpty(shape.ImageData.ImageBytes);
            }

            TestUtil.VerifyWebResponseStatusCode(HttpStatusCode.OK, "http://www.google.com/images/logos/ps_logo2.png");
            TestUtil.VerifyWebResponseStatusCode(HttpStatusCode.OK, AsposeLogoUrl);
        }
#endif
    }
}