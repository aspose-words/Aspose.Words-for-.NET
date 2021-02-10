#if NET462
using System.Drawing;
using System.Drawing.Imaging;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using NUnit.Framework;
using System.IO;

namespace DocsExamples.Rendering_and_Printing
{
    internal class RenderingShapes : DocsExamplesBase
    {
        [Test]
        public void RenderShapeAsEmf()
        {
            Document doc = new Document(MyDir + "Rendering.docx");
            
            // Retrieve the target shape from the document.
            Shape shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);

            //ExStart:RenderShapeAsEmf
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
            
            Shape shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);

            //ExStart:RenderShapeAsJpeg
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

        [Test]
        //ExStart:RenderShapeToGraphics
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
        public void RenderCellToImage()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            //ExStart:RenderCellToImage
            Cell cell = (Cell)doc.GetChild(NodeType.Cell, 2, true);
            RenderNode(cell, ArtifactsDir + "RenderShape.RenderCellToImage.png", null);
            //ExEnd:RenderCellToImage
        }

        [Test]
        public void RenderRowToImage()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            //ExStart:RenderRowToImage
            Row row = (Row) doc.GetChild(NodeType.Row, 0, true);
            RenderNode(row, ArtifactsDir + "RenderShape.RenderRowToImage.png", null);
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

            RenderNode(textBoxShape.LastParagraph, ArtifactsDir + "RenderShape.RenderParagraphToImage.png", options);
            //ExEnd:RenderParagraphToImage
        }

        [Test]
        public void FindShapeSizes()
        {
            Document doc = new Document(MyDir + "Rendering.docx");
            
            Shape shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);

            //ExStart:FindShapeSizes
            Size shapeRenderedSize = shape.GetShapeRenderer().GetSizeInPixels(1.0f, 96.0f);

            using (Bitmap image = new Bitmap(shapeRenderedSize.Width, shapeRenderedSize.Height))
            {
                using (Graphics graphics = Graphics.FromImage(image))
                {
                    // Render shape onto the graphics object using the RenderToScale or RenderToSize methods of ShapeRenderer class.
                }
            }
            //ExEnd:FindShapeSizes
        }

        [Test]
        public void RenderShapeImage()
        {
            Document doc = new Document(MyDir + "Rendering.docx");
            
            Shape shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);

            //ExStart:RenderShapeImage
            shape.GetShapeRenderer().Save(ArtifactsDir + "RenderShape.RenderShapeImage.jpg", null);
            //ExEnd:RenderShapeImage
        }

        /// <summary>
        /// Renders any node in a document to the path specified using the image save options.
        /// </summary>
        /// <param name="node">The node to render.</param>
        /// <param name="filePath">The path to save the rendered image to.</param>
        /// <param name="imageOptions">The image options to use during rendering. This can be null.</param>
        public void RenderNode(Node node, string filePath, ImageSaveOptions imageOptions)
        {
            if (imageOptions == null)
                imageOptions = new ImageSaveOptions(FileFormatUtil.ExtensionToSaveFormat(Path.GetExtension(filePath)));

            // Store the paper color to be used on the final image and change to transparent.
            // This will cause any content around the rendered node to be removed later on.
            Color savePaperColor = imageOptions.PaperColor;
            imageOptions.PaperColor = Color.Transparent;

            // There a bug which affects the cache of a cloned node.
            // To avoid this, we clone the entire document, including all nodes,
            // finding the matching node in the cloned document and rendering that instead.
            Document doc = (Document) node.Document.Clone(true);
            node = doc.GetChild(NodeType.Any, node.Document.GetChildNodes(NodeType.Any, true).IndexOf(node), true);

            // Create a temporary shape to store the target node in. This shape will be rendered to retrieve
            // the rendered content of the node.
            Shape shape = new Shape(doc, ShapeType.TextBox);
            Section parentSection = (Section) node.GetAncestor(NodeType.Section);

            // Assume that the node cannot be larger than the page in size.
            shape.Width = parentSection.PageSetup.PageWidth;
            shape.Height = parentSection.PageSetup.PageHeight;
            shape.FillColor = Color.Transparent;

            // Don't draw a surronding line on the shape.
            shape.Stroked = false;

            // Move up through the DOM until we find a suitable node to insert into a Shape
            // (a node with a parent can contain paragraphs, tables the same as a shape). Each parent node is cloned
            // on the way up so even a descendant node passed to this method can be rendered. Since we are working
            // with the actual nodes of the document we need to clone the target node into the temporary shape.
            Node currentNode = node;
            while (!(currentNode.ParentNode is InlineStory || currentNode.ParentNode is Story ||
                     currentNode.ParentNode is ShapeBase))
            {
                CompositeNode parent = (CompositeNode) currentNode.ParentNode.Clone(false);
                currentNode = currentNode.ParentNode;
                parent.AppendChild(node.Clone(true));
                node = parent; // Store this new node to be inserted into the shape.
            }

            // We must add the shape to the document tree to have it rendered.
            shape.AppendChild(node.Clone(true));
            parentSection.Body.FirstParagraph.AppendChild(shape);

            // Render the shape to stream so we can take advantage of the effects of the ImageSaveOptions class.
            // Retrieve the rendered image and remove the shape from the document.
            MemoryStream stream = new MemoryStream();
            ShapeRenderer renderer = shape.GetShapeRenderer();
            renderer.Save(stream, imageOptions);
            shape.Remove();

            Rectangle crop = renderer.GetOpaqueBoundsInPixels(imageOptions.Scale, imageOptions.HorizontalResolution,
                imageOptions.VerticalResolution);

            using (Bitmap renderedImage = new Bitmap(stream))
            {
                Bitmap croppedImage = new Bitmap(crop.Width, crop.Height);
                croppedImage.SetResolution(imageOptions.HorizontalResolution, imageOptions.VerticalResolution);

                // Create the final image with the proper background color.
                using (Graphics g = Graphics.FromImage(croppedImage))
                {
                    g.Clear(savePaperColor);
                    g.DrawImage(renderedImage, new Rectangle(0, 0, croppedImage.Width, croppedImage.Height), crop.X,
                        crop.Y, crop.Width, crop.Height, GraphicsUnit.Pixel);

                    croppedImage.Save(filePath);
                }
            }
        }

        /// <summary>
        /// Finds the minimum bounding box around non-transparent pixels in a Bitmap.
        /// </summary>
        public Rectangle FindBoundingBoxAroundNode(Bitmap originalBitmap)
        {
            Point min = new Point(int.MaxValue, int.MaxValue);
            Point max = new Point(int.MinValue, int.MinValue);

            for (int x = 0; x < originalBitmap.Width; ++x)
            {
                for (int y = 0; y < originalBitmap.Height; ++y)
                {
                    // Note that you can speed up this part of the algorithm using LockBits and unsafe code instead of GetPixel.
                    Color pixelColor = originalBitmap.GetPixel(x, y);

                    // For each pixel that is not transparent, calculate the bounding box around it.
                    if (pixelColor.ToArgb() != Color.Empty.ToArgb())
                    {
                        min.X = System.Math.Min(x, min.X);
                        min.Y = System.Math.Min(y, min.Y);
                        max.X = System.Math.Max(x, max.X);
                        max.Y = System.Math.Max(y, max.Y);
                    }
                }
            }

            // Add one pixel to the width and height to avoid clipping.
            return new Rectangle(min.X, min.Y, max.X - min.X + 1, max.Y - min.Y + 1);
        }
    }
}
#endif