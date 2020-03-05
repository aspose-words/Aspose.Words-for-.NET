// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
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
#if NETFRAMEWORK || JAVA
using Aspose.Words.Loading;
using System.Net;
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
            // DocumentBase is the abstract base class for the Document and GlossaryDocument classes
            Document doc = new Document();

            GlossaryDocument glossaryDoc = new GlossaryDocument();
            doc.GlossaryDocument = glossaryDoc;
            //ExEnd
        }

        [Test]
        public void SetPageColor()
        {
            //ExStart
            //ExFor:DocumentBase.PageColor
            //ExSummary:Shows how to set the page color.
            Document doc = new Document();

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
            //ExSummary:Shows how to import node from source document to destination document.
            Document src = new Document();
            Document dst = new Document();

            // Add text to both documents
            src.FirstSection.Body.FirstParagraph.AppendChild(new Run(src, "Source document first paragraph text."));
            dst.FirstSection.Body.FirstParagraph.AppendChild(new Run(dst, "Destination document first paragraph text."));

            // In order for a child node to be successfully appended to another node in a document,
            // both nodes must have the same parent document, or an exception is thrown
            Assert.AreNotEqual(dst, src.FirstSection.Document);
            Assert.Throws<ArgumentException>(() => { dst.AppendChild(src.FirstSection); });

            // For that reason, we can't just append a section of the source document to the destination document using Node.AppendChild()
            // Document.ImportNode() lets us get around this by creating a clone of a node and sets its parent to the calling document
            Section importedSection = (Section)dst.ImportNode(src.FirstSection, true);

            // Now it is ready to be placed in the document
            dst.AppendChild(importedSection);

            // Our document now contains both the original and imported section
            Assert.AreEqual("Destination document first paragraph text.\r\nSource document first paragraph text.\r\n",
                dst.ToString(SaveFormat.Text));
            //ExEnd

            Assert.AreNotEqual(importedSection, src.FirstSection);
            Assert.AreNotEqual(importedSection.Document, src.FirstSection.Document);
            Assert.AreEqual(importedSection.Body.FirstParagraph.GetText(),
                src.FirstSection.Body.FirstParagraph.GetText());
        }

        [Test]
        public void ImportNodeCustom()
        {
            //ExStart
            //ExFor:DocumentBase.ImportNode(Node, System.Boolean, ImportFormatMode)
            //ExSummary:Shows how to import node from source document to destination document with specific options.
            // Create two documents with two styles that differ in font but have the same name
            Document src = new Document();
            Style srcStyle = src.Styles.Add(StyleType.Character, "My style");
            srcStyle.Font.Name = "Courier New";
            DocumentBuilder srcBuilder = new DocumentBuilder(src);
            srcBuilder.Font.Style = srcStyle;
            srcBuilder.Writeln("Source document text.");

            Document dst = new Document();
            Style dstStyle = dst.Styles.Add(StyleType.Character, "My style");
            dstStyle.Font.Name = "Calibri";
            DocumentBuilder dstBuilder = new DocumentBuilder(dst);
            dstBuilder.Font.Style = dstStyle;
            dstBuilder.Writeln("Destination document text.");

            // Import the Section from the destination document into the source document, causing a style name collision
            // If we use destination styles then the imported source text with the same style name as destination text
            // will adopt the destination style 
            Section importedSection = (Section)dst.ImportNode(src.FirstSection, true, ImportFormatMode.UseDestinationStyles);
            Assert.AreEqual("Source document text.", importedSection.Body.Paragraphs[0].Runs[0].GetText().Trim()); //ExSkip
            Assert.IsNull(dst.Styles["My style_0"]); //ExSkip
            Assert.AreEqual(dstStyle.Font.Name, importedSection.Body.FirstParagraph.Runs[0].Font.Name);
            Assert.AreEqual(dstStyle.Name, importedSection.Body.FirstParagraph.Runs[0].Font.StyleName);

            // If we use ImportFormatMode.KeepDifferentStyles,
            // the source style is preserved and the naming clash is resolved by adding a suffix 
            dst.ImportNode(src.FirstSection, true, ImportFormatMode.KeepDifferentStyles);
            Assert.AreEqual(dstStyle.Font.Name, dst.Styles["My style"].Font.Name);
            Assert.AreEqual(srcStyle.Font.Name, dst.Styles["My style_0"].Font.Name);
            //ExEnd
        }

        [Test]
        public void BackgroundShape()
        {
            //ExStart
            //ExFor:DocumentBase.BackgroundShape
            //ExSummary:Shows how to set the background shape of a document.
            Document doc = new Document();
            Assert.IsNull(doc.BackgroundShape);

            // A background shape can only be a rectangle
            // We will set the color of this rectangle to light blue
            Shape shapeRectangle = new Shape(doc, ShapeType.Rectangle);
            doc.BackgroundShape = shapeRectangle;

            // This rectangle covers the entire page in the output document
            // We can also do this by setting doc.PageColor
            shapeRectangle.FillColor = System.Drawing.Color.LightBlue;
            doc.Save(ArtifactsDir + "DocumentBase.BackgroundShapeFlatColor.docx");

            // Setting the image will override the flat background color with the image
            shapeRectangle.ImageData.SetImage(ImageDir + "Transparent background logo.png");
            Assert.IsTrue(doc.BackgroundShape.HasImage);

            // This image is a photo with a white background
            // To make it suitable as a watermark, we will need to do some image processing
            // The default values for these variables are 0.5, so here we are lowering the contrast and increasing the brightness
            shapeRectangle.ImageData.Contrast = 0.2;
            shapeRectangle.ImageData.Brightness = 0.7;

            // Microsoft Word does not support images in background shapes, so even though we set the background as an image,
            // the output will show a light blue background like before
            // However, we can see our watermark in an output pdf
            doc.Save(ArtifactsDir + "DocumentBase.BackgroundShape.pdf");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBase.BackgroundShapeFlatColor.docx");
            Assert.AreEqual(System.Drawing.Color.LightBlue.ToArgb(), doc.BackgroundShape.FillColor.ToArgb());
        }

        #if NETFRAMEWORK || JAVA
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
        //ExSummary:Shows how to process inserted resources differently.
        [Test] //ExSkip
        public void ResourceLoadingCallback()
        {
            Document doc = new Document();

            // Enable our custom image loading
            doc.ResourceLoadingCallback = new ImageNameHandler();

            DocumentBuilder builder = new DocumentBuilder(doc);

            // We usually insert images as a uri or byte array, but there are many other possibilities with ResourceLoadingCallback
            // In this case we are referencing images with simple names and keep the image fetching logic somewhere else
            builder.InsertImage("Google Logo");
            builder.InsertImage("Aspose Logo");
            builder.InsertImage("My Watermark");

            // Images belong to Shape objects, which are placed and scaled in the document
            Assert.AreEqual(3, doc.GetChildNodes(NodeType.Shape, true).Count);

            doc.Save(ArtifactsDir + "DocumentBase.ResourceLoadingCallback.docx");
            TestResourceLoadingCallback(new Document(ArtifactsDir + "DocumentBase.ResourceLoadingCallback.docx")); //ExSkip
        }

        private class ImageNameHandler : IResourceLoadingCallback
        {
            public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
            {
                if (args.ResourceType == ResourceType.Image)
                {
                    // builder.InsertImage expects a uri so inputs like "Google Logo" would normally trigger a FileNotFoundException
                    // We can still process those inputs and find an image any way we like, as long as an image byte array is passed to args.SetData()
                    if (args.OriginalUri == "Google Logo")
                    {
                        using (WebClient webClient = new WebClient())
                        {
                            byte[] imageBytes =
                                webClient.DownloadData("http://www.google.com/images/logos/ps_logo2.png");
                            args.SetData(imageBytes);
                            // We need this return statement any time a resource is loaded in a custom manner
                            return ResourceLoadingAction.UserProvided;
                        }
                    }

                    if (args.OriginalUri == "Aspose Logo")
                    {
                        using (WebClient webClient = new WebClient())
                        {
                            byte[] imageBytes = webClient.DownloadData(AsposeLogoUrl);
                            args.SetData(imageBytes);
                            return ResourceLoadingAction.UserProvided;
                        }
                    }

                    // We can find and add an image any way we like, as long as args.SetData() is called with some image byte array as a parameter
                    if (args.OriginalUri == "My Watermark")
                    {
                        System.Drawing.Image watermark = System.Drawing.Image.FromFile(ImageDir + "Transparent background logo.png");

                        System.Drawing.ImageConverter converter = new System.Drawing.ImageConverter();
                        byte[] imageBytes = (byte[])converter.ConvertTo(watermark, typeof(byte[]));
                        args.SetData(imageBytes);

                        return ResourceLoadingAction.UserProvided;
                    }
                }

                // All other resources such as documents, CSS stylesheets and images passed as uris are handled as they were normally
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
        }
        #endif
    }
}