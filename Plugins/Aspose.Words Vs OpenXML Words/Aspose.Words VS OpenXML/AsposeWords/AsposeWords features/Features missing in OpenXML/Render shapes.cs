// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Layout;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features.Features_missing_in_OpenXML
{
    [TestFixture]
    public class RenderShapes : TestUtil
    {
        [Test]
        public void RenderShapeAsEmf()
        {
            Document doc = new Document(MyDir + "Rendering.docx");
            // Retrieve the target shape from the document.
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            //ExStart:RenderShapeAsEmf
            //GistId:7fc867ac8ef1b729b6f70580fbc5b3f9
            ShapeRenderer render = shape.GetShapeRenderer();
            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Emf)
            {
                Scale = 1.5f
            };

            render.Save(ArtifactsDir + "RenderShape.RenderShapeAsEmf.emf", imageOptions);
            //ExEnd:RenderShapeAsEmf
        }

        [Test]
        public void RenderShapeAsJpeg()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            //ExStart:RenderShapeAsJpeg
            //GistId:7fc867ac8ef1b729b6f70580fbc5b3f9
            ShapeRenderer render = new ShapeRenderer(shape);
            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Output the image in gray scale
                ImageColorMode = ImageColorMode.Grayscale,
                // Reduce the brightness a bit (default is 0.5f)
                ImageBrightness = 0.45f
            };

            using (FileStream stream = new FileStream(ArtifactsDir + "RenderShape.RenderShapeAsJpeg.jpg", FileMode.Create))
            {
                render.Save(stream, imageOptions);
            }
            //ExEnd:RenderShapeAsJpeg
        }
#if NET48 || JAVA
        [Test]
        //ExStart:RenderShapeToGraphics
        //GistId:7fc867ac8ef1b729b6f70580fbc5b3f9
        public void RenderShapeToGraphics()
        {
            Document doc = new Document(MyDir + "Rendering.docx");
            
            Shape shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            ShapeRenderer render = shape.GetShapeRenderer();

            // Find the size that the shape will be rendered to at the specified scale and resolution.
            Size shapeSizeInPixels = render.GetSizeInPixels(1.0f, 96.0f);
            // Rotating the shape may result in clipping as the image canvas is too small. Find the longest side
            // and make sure that the graphics canvas is large enough to compensate for this.
            int maxSide = System.Math.Max(shapeSizeInPixels.Width, shapeSizeInPixels.Height);

            using (Bitmap image = new Bitmap((int) (maxSide * 1.25), (int) (maxSide * 1.25)))
            {
                // Rendering to a graphics object means we can specify settings and transformations to be applied to the rendered shape.
                // In our case we will rotate the rendered shape.
                using (Graphics graphics = Graphics.FromImage(image))
                {
                    // Clear the shape with the background color of the document.
                    graphics.Clear(shape.Document.PageColor);
                    // Center the rotation using the translation method below.
                    graphics.TranslateTransform((float) image.Width / 8, (float) image.Height / 2);
                    // Rotate the image by 45 degrees.
                    graphics.RotateTransform(45);
                    // Undo the translation.
                    graphics.TranslateTransform(-(float) image.Width / 8, -(float) image.Height / 2);

                    // Render the shape onto the graphics object.
                    render.RenderToSize(graphics, 0, 0, shapeSizeInPixels.Width, shapeSizeInPixels.Height);
                }

                image.Save(ArtifactsDir + "RenderShape.RenderShapeToGraphics.png", ImageFormat.Png);
            }
        }
        //ExEnd:RenderShapeToGraphics

        [Test]
        public void FindShapeSizes()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            //ExStart:FindShapeSizes
            //GistId:7fc867ac8ef1b729b6f70580fbc5b3f9
            Size shapeRenderedSize = shape.GetShapeRenderer().GetSizeInPixels(1.0f, 96.0f);

            using (Bitmap image = new Bitmap(shapeRenderedSize.Width, shapeRenderedSize.Height))
            {
                using (Graphics graphics = Graphics.FromImage(image))
                {
                    // Render shape onto the graphics object using the RenderToScale
                    // or RenderToSize methods of ShapeRenderer class.
                }
            }
            //ExEnd:FindShapeSizes
        }
#endif

        [Test]
        public void RenderCellToImage()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            //ExStart:RenderCellToImage
            Cell cell = (Cell)doc.GetChild(NodeType.Cell, 2, true);
            Document tmp = ConvertToImage(doc, cell);
            tmp.Save(ArtifactsDir + "RenderShape.RenderCellToImage.png");
            //ExEnd:RenderCellToImage
        }

        [Test]
        public void RenderRowToImage()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            //ExStart:RenderRowToImage
            Row row = (Row)doc.GetChild(NodeType.Row, 0, true);
            Document tmp = ConvertToImage(doc, row);
            tmp.Save(ArtifactsDir + "RenderShape.RenderRowToImage.png");
            //ExEnd:RenderRowToImage
        }

        [Test]
        public void RenderParagraphToImage()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            //ExStart:RenderParagraphToImage
            Shape textBoxShape = builder.InsertShape(ShapeType.TextBox, 150, 100);

            builder.MoveTo(textBoxShape.LastParagraph);
            builder.Write("Vertical text");

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                PaperColor = Color.LightPink
            };

            Document tmp = ConvertToImage(doc, textBoxShape.LastParagraph);
            tmp.Save(ArtifactsDir + "RenderShape.RenderParagraphToImage.png");
            //ExEnd:RenderParagraphToImage
        }

        [Test]
        public void RenderShapeImage()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            //ExStart:RenderShapeImage
            //GistId:7fc867ac8ef1b729b6f70580fbc5b3f9
            shape.GetShapeRenderer().Save(ArtifactsDir + "RenderShape.RenderShapeImage.jpg", new ImageSaveOptions(SaveFormat.Jpeg));
            //ExEnd:RenderShapeImage
        }

        /// <summary>
        /// Renders any node in a document into an image.
        /// </summary>
        /// <param name="doc">The current document.</param>
        /// <param name="node">The node to render.</param>
        private static Document ConvertToImage(Document doc, CompositeNode node)
        {
            Document tmp = CreateTemporaryDocument(doc, node);
            AppendNodeContent(tmp, node);
            AdjustDocumentLayout(tmp);
            return tmp;
        }

        /// <summary>
        /// Creates a temporary document for further rendering.
        /// </summary>
        private static Document CreateTemporaryDocument(Document doc, CompositeNode node)
        {
            Document tmp = (Document)doc.Clone(false);
            tmp.Sections.Add(tmp.ImportNode(node.GetAncestor(NodeType.Section), false, ImportFormatMode.UseDestinationStyles));
            tmp.FirstSection.AppendChild(new Body(tmp));
            tmp.FirstSection.PageSetup.TopMargin = 0;
            tmp.FirstSection.PageSetup.BottomMargin = 0;

            return tmp;
        }

        /// <summary>
        /// Adds a node to a temporary document.
        /// </summary>
        private static void AppendNodeContent(Document tmp, CompositeNode node)
        {
            if (node is HeaderFooter headerFooter)
                foreach (Node hfNode in headerFooter.GetChildNodes(NodeType.Any, false))
                    tmp.FirstSection.Body.AppendChild(tmp.ImportNode(hfNode, true, ImportFormatMode.UseDestinationStyles));
            else
                AppendNonHeaderFooterContent(tmp, node);
        }

        private static void AppendNonHeaderFooterContent(Document tmp, CompositeNode node)
        {
            Node parentNode = node.ParentNode;
            while (!(parentNode is InlineStory || parentNode is Story || parentNode is ShapeBase))
            {
                CompositeNode parent = (CompositeNode)parentNode.Clone(false);
                parent.AppendChild(node.Clone(true));
                node = parent;

                parentNode = parentNode.ParentNode;
            }

            tmp.FirstSection.Body.AppendChild(tmp.ImportNode(node, true, ImportFormatMode.UseDestinationStyles));
        }

        /// <summary>
        /// Adjusts the layout of the document to fit the content area.
        /// </summary>
        private static void AdjustDocumentLayout(Document tmp)
        {
            LayoutEnumerator enumerator = new LayoutEnumerator(tmp);
            RectangleF rect = RectangleF.Empty;
            rect = CalculateVisibleRect(enumerator, rect);

            tmp.FirstSection.PageSetup.PageHeight = rect.Height;
            tmp.UpdatePageLayout();
        }

        /// <summary>
        /// Calculates the visible area of the content.
        /// </summary>
        private static RectangleF CalculateVisibleRect(LayoutEnumerator enumerator, RectangleF rect)
        {
            RectangleF result = rect;
            do
            {
                if (enumerator.MoveFirstChild())
                {
                    if (enumerator.Type == LayoutEntityType.Line || enumerator.Type == LayoutEntityType.Span)
                        result = result.IsEmpty ? enumerator.Rectangle : RectangleF.Union(result, enumerator.Rectangle);
                    result = CalculateVisibleRect(enumerator, result);
                    enumerator.MoveParent();
                }
            } while (enumerator.MoveNext());

            return result;
        }
    }
}

