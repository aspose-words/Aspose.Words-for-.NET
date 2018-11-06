using System.Net;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using NUnit.Framework;

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

            doc.Save(MyDir + @"\Artifacts\DocumentBase.SetPageColor.docx");
            //ExEnd
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
            dst.FirstSection.Body.FirstParagraph.AppendChild(new Run(dst,
                "Destination document first paragraph text."));

            // If we want to add the section from doc2 to doc1, we can't just append them like this:
            // dst.AppendChild(src.FirstSection);
            // Uncommenting that line throws an exception because doc2's first section belongs to doc2,
            // but each node in a document must belong to the document
            Assert.AreNotEqual(dst, src.FirstSection.Document);

            // We can create a new node that belongs to the destination document
            Section importedSection = (Section) dst.ImportNode(src.FirstSection, true);

            // It has the same content but it is not the same node nor do they have the same owner
            Assert.AreNotEqual(importedSection, src.FirstSection);
            Assert.AreNotEqual(importedSection.Document, src.FirstSection.Document);
            Assert.AreEqual(importedSection.Body.FirstParagraph.GetText(),
                src.FirstSection.Body.FirstParagraph.GetText());

            // Now it is ready to be placed in the document
            dst.AppendChild(importedSection);

            // Our document does indeed contain both the original and imported section
            Assert.AreEqual("Destination document first paragraph text.\r\nSource document first paragraph text.\r\n",
                dst.ToString(SaveFormat.Text));
            //ExEnd
        }

        [Test]
        public void ImportNodeCustom()
        {
            //ExStart
            //ExFor:DocumentBase.ImportNode(Node, System.Boolean, ImportFormatMode)
            //ExSummary:Shows how to import node from source document to destination document with specific options.
            // Create two documents with two styles that aren't the same but have the same name
            Document src = new Document();
            Style srcStyle = src.Styles.Add(StyleType.Character, "My style");
            DocumentBuilder srcBuilder = new DocumentBuilder(src);
            srcBuilder.Font.Style = srcStyle;
            srcBuilder.Writeln("Source document text.");

            Document dst = new Document();
            Style dstStyle = dst.Styles.Add(StyleType.Character, "My style");
            dstStyle.Font.Bold = true;
            DocumentBuilder dstBuilder = new DocumentBuilder(dst);
            dstBuilder.Font.Style = dstStyle;
            srcBuilder.Writeln("Destination document text.");

            dst.ImportNode(src.FirstSection, true, ImportFormatMode.UseDestinationStyles);

            Assert.IsNull(dst.Styles["My style_0"]);

            dst.ImportNode(src.FirstSection, true, ImportFormatMode.KeepDifferentStyles);

            Assert.IsNotNull(dst.Styles["My style_0"]);
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
            // We will set the colour of this rectangle to light blue
            Shape shapeRectangle = new Shape(doc, ShapeType.Rectangle);
            doc.BackgroundShape = shapeRectangle;

            // This rectangle covers the entire page in the output document
            // We can also do this by setting doc.PageColor
            shapeRectangle.FillColor = System.Drawing.Color.LightBlue;
            doc.Save(MyDir + @"\Artifacts\DocumentBase.BackgroundShapeFlatColor.docx");

            // Setting the image will override the flat background colour with the image
            shapeRectangle.ImageData.SetImage(MyDir + @"\Images\Watermark.png");
            Assert.IsTrue(doc.BackgroundShape.HasImage);

            // This image is a photo with a white background
            // To make it suitable as a watermark, we will need to do some image processing
            // The default values for these variables are 0.5, so here we are lowering the contrast and increasing the brightness
            shapeRectangle.ImageData.Contrast = 0.2;
            shapeRectangle.ImageData.Brightness = 0.7;

            // Microsoft Word does not support images in background shapes, so even though we set the background as an image,
            // the output will show a light blue background like before
            // However, we can see our watermark in an output pdf
            doc.Save(MyDir + @"\Artifacts\DocumentBase.BackgroundShapeWatermark.pdf");
            //ExEnd
        }

#if !(NETSTANDARD2_0 || __MOBILE__)
        //ExStart
        //ExFor:DocumentBase.ResourceLoadingCallback
        //ExSummary:Shows how to process inserted resources differently.
        [Test] //ExSkip
        public void ResourceLoadingCallback()
        {
            Document doc = new Document();

            // Images belong to NodeType.Shape
            // There are none in a blank document
            Assert.AreEqual(0, doc.GetChildNodes(NodeType.Shape, true).Count);

            // Enable our custom image loading
            doc.ResourceLoadingCallback = new ImageNameHandler();

            DocumentBuilder builder = new DocumentBuilder(doc);

            // We usually insert images as a uri or byte array, but there are many other possibilities with ResourceLoadingCallback
            // In this case we are referencing images with simple names and keep the image fetching logic somewhere else
            builder.InsertImage("Google Logo");
            builder.InsertImage("Aspose Logo");
            builder.InsertImage("My Watermark");

            Assert.AreEqual(3, doc.GetChildNodes(NodeType.Shape, true).Count);

            doc.Save(MyDir + @"\Artifacts\DocumentBase.ResourceLoadingCallback.docx");
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
                            byte[] imageBytes = webClient.DownloadData("https://www.aspose.com/Images/aspose-logo.jpg");
                            args.SetData(imageBytes);
                            return ResourceLoadingAction.UserProvided;
                        }
                    }

                    // We can find and add an image any way we like, as long as args.SetData() is called with some image byte array as a parameter
                    if (args.OriginalUri == "My Watermark")
                    {
                        System.Drawing.Image watermark =
                            System.Drawing.Image.FromFile(MyDir + @"\Images\Watermark.png");
                        System.Drawing.ImageConverter converter = new System.Drawing.ImageConverter();
                        byte[] imageBytes = (byte[]) converter.ConvertTo(watermark, typeof(byte[]));
                        args.SetData(imageBytes);

                        return ResourceLoadingAction.UserProvided;
                    }
                }

                // All other resources such as documents, CSS stylesheets and images passed as uris are handled as they were normally
                return ResourceLoadingAction.Default;
            }
        }
        //ExEnd
#endif
    }
}