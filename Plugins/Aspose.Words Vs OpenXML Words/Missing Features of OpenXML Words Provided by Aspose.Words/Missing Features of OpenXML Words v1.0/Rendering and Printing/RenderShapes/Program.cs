// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;
using Aspose.Words.Tables;


namespace RenderShapes
{
    class Program
    {
        static void Main(string[] args)
        {
            // Sample infrastructure.
            string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar;
            string dataDir = new Uri(new Uri(exeDir), @"../../Data/").LocalPath;

            // Load the documents which store the shapes we want to render.
            Document doc = new Document(dataDir + "RenderShapes.doc");
            Document doc2 = new Document(dataDir + "RenderShapes.docx");

            // Retrieve the target shape from the document. In our sample document this is the first shape.
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            DrawingML drawingML = (DrawingML)doc2.GetChild(NodeType.DrawingML, 0, true);

            // Test rendering of different types of nodes.
            RenderShapeToDisk(dataDir, shape);
            RenderShapeToStream(dataDir, shape);
            RenderShapeToGraphics(dataDir, shape);
            RenderDrawingMLToDisk(dataDir, drawingML);
            RenderCellToImage(dataDir, doc);
            RenderRowToImage(dataDir, doc);
            RenderParagraphToImage(dataDir, doc);
            FindShapeSizes(shape);
        }

        public static void RenderShapeToDisk(string dataDir, Shape shape)
        {
            //ExStart
            //ExFor:ShapeRenderer
            //ExFor:ShapeBase.GetShapeRenderer
            //ExFor:ImageSaveOptions
            //ExFor:ImageSaveOptions.Scale
            //ExFor:ShapeRenderer.Save(String, ImageSaveOptions)
            //ExId:RenderShapeToDisk
            //ExSummary:Shows how to render a shape independent of the document to an EMF image and save it to disk.
            // The shape render is retrieved using this method. This is made into a separate object from the shape as it internally
            // caches the rendered shape.
            ShapeRenderer r = shape.GetShapeRenderer();

            // Define custom options which control how the image is rendered. Render the shape to the JPEG raster format.
            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Emf)
            {
                Scale = 1.5f
            };

            // Save the rendered image to disk.
            r.Save(dataDir + "TestFile.RenderToDisk Out.emf", imageOptions);
            //ExEnd
        }

        public static void RenderShapeToStream(string dataDir, Shape shape)
        {
            //ExStart
            //ExFor:ShapeRenderer
            //ExFor:ShapeRenderer.#ctor(ShapeBase)
            //ExFor:ImageSaveOptions.ImageColorMode
            //ExFor:ImageSaveOptions.ImageBrightness
            //ExFor:ShapeRenderer.Save(Stream, ImageSaveOptions)
            //ExId:RenderShapeToStream
            //ExSummary:Shows how to render a shape independent of the document to a JPEG image and save it to a stream.
            // We can also retrieve the renderer for a shape by using the ShapeRenderer constructor.
            ShapeRenderer r = new ShapeRenderer(shape);

            // Define custom options which control how the image is rendered. Render the shape to the vector format EMF.
            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Output the image in gray scale
                ImageColorMode = ImageColorMode.Grayscale,

                // Reduce the brightness a bit (default is 0.5f).
                ImageBrightness = 0.45f
            };

            FileStream stream = new FileStream(dataDir + "TestFile.RenderToStream Out.jpg", FileMode.CreateNew);

            // Save the rendered image to the stream using different options.
            r.Save(stream, imageOptions);
            //ExEnd
        }

        public static void RenderDrawingMLToDisk(string dataDir, DrawingML drawingML)
        {
            //ExStart
            //ExFor:DrawingML.GetShapeRenderer
            //ExFor:ShapeRenderer.Save(String, ImageSaveOptions)
            //ExFor:DrawingML
            //ExId:RenderDrawingMLToDisk
            //ExSummary:Shows how to render a DrawingML image independent of the document to a JPEG image on the disk.
            // Save the DrawingML image to disk in JPEG format and using default options.
            drawingML.GetShapeRenderer().Save(dataDir + "TestFile.RenderDrawingML Out.jpg", null);
            //ExEnd
        }

        public static void RenderShapeToGraphics(string dataDir, Shape shape)
        {
            //ExStart
            //ExFor:ShapeRenderer
            //ExFor:ShapeBase.GetShapeRenderer
            //ExFor:ShapeRenderer.GetSizeInPixels
            //ExFor:ShapeRenderer.RenderToSize
            //ExId:RenderShapeToGraphics
            //ExSummary:Shows how to render a shape independent of the document to a .NET Graphics object and apply rotation to the rendered image.
            // The shape renderer is retrieved using this method. This is made into a separate object from the shape as it internally
            // caches the rendered shape.
            ShapeRenderer r = shape.GetShapeRenderer();

            // Find the size that the shape will be rendered to at the specified scale and resolution.
            Size shapeSizeInPixels = r.GetSizeInPixels(1.0f, 96.0f);

            // Rotating the shape may result in clipping as the image canvas is too small. Find the longest side
            // and make sure that the graphics canvas is large enough to compensate for this.
            int maxSide = Math.Max(shapeSizeInPixels.Width, shapeSizeInPixels.Height);

            using (Bitmap image = new Bitmap((int)(maxSide * 1.25), (int)(maxSide * 1.25)))
            {
                // Rendering to a graphics object means we can specify settings and transformations to be applied to 
                // the shape that is rendered. In our case we will rotate the rendered shape.
                using (Graphics gr = Graphics.FromImage(image))
                {
                    // Clear the shape with the background color of the document.
                    gr.Clear(Color.White);
                    // Center the rotation using translation method below
                    gr.TranslateTransform((float)image.Width / 8, (float)image.Height / 2);
                    // Rotate the image by 45 degrees.
                    gr.RotateTransform(45);
                    // Undo the translation.
                    gr.TranslateTransform(-(float)image.Width / 8, -(float)image.Height / 2);

                    // Render the shape onto the graphics object.
                    r.RenderToSize(gr, 0, 0, shapeSizeInPixels.Width, shapeSizeInPixels.Height);
                }

                image.Save(dataDir + "TestFile.RenderToGraphics.png", ImageFormat.Png);
            }
            //ExEnd
        }

        public static void RenderCellToImage(string dataDir, Document doc)
        {
            //ExStart
            //ExId:RenderCellToImage
            //ExSummary:Shows how to render a cell of a table independent of the document.
            Cell cell = (Cell)doc.GetChild(NodeType.Cell, 2, true); // The third cell in the first table.
            RenderNode(cell, dataDir + "TestFile.RenderCell Out.png", null);
            //ExEnd
        }

        public static void RenderRowToImage(string dataDir, Document doc)
        {
            //ExStart
            //ExId:RenderRowToImage
            //ExSummary:Shows how to render a row of a table independent of the document.
            Row row = (Row)doc.GetChild(NodeType.Row, 0, true); // The first row in the first table.
            RenderNode(row, dataDir + "TestFile.RenderRow Out.png", null);
            //ExEnd
        }

        public static void RenderParagraphToImage(string dataDir, Document doc)
        {
            //ExStart
            //ExFor:Shape.LastParagraph
            //ExId:RenderParagraphToImage
            //ExSummary:Shows how to render a paragraph with a custom background color independent of the document. 
            // Retrieve the first paragraph in the main shape.
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Paragraph paragraph = (Paragraph)shape.LastParagraph;

            // Save the node with a light pink background.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
            options.PaperColor = Color.LightPink;

            RenderNode(paragraph, dataDir + "TestFile.RenderParagraph Out.png", options);
            //ExEnd
        }

        public static void FindShapeSizes(Shape shape)
        {
            //ExStart
            //ExFor:ShapeRenderer.SizeInPoints
            //ExId:ShapeRendererSizeInPoints
            //ExSummary:Demonstrates how to find the size of a shape in the document and the size of the shape when rendered.
            SizeF shapeSizeInDocument = shape.GetShapeRenderer().SizeInPoints;
            float width = shapeSizeInDocument.Width; // The width of the shape.
            float height = shapeSizeInDocument.Height; // The height of the shape.
            //ExEnd

            //ExStart
            //ExFor:ShapeRenderer.GetSizeInPixels
            //ExId:ShapeRendererGetSizeInPixels
            //ExSummary:Shows how to create a new Bitmap and Graphics object with the width and height of the shape to be rendered.
            // We will render the shape at normal size and 96dpi. Calculate the size in pixels that the shape will be rendered at.
            Size shapeRenderedSize = shape.GetShapeRenderer().GetSizeInPixels(1.0f, 96.0f);

            using (Bitmap image = new Bitmap(shapeRenderedSize.Width, shapeRenderedSize.Height))
            {
                using (Graphics g = Graphics.FromImage(image))
                {
                    // Render shape onto the graphics object using the RenderToScale or RenderToSize methods of ShapeRenderer class.
                }
            }
            //ExEnd
        }

        //ExStart
        //ExId:RenderNode
        //ExSummary:Shows how to render a node independent of the document by building on the functionality provided by ShapeRenderer class.
        /// <summary>
        /// Renders any node in a document to the path specified using the image save options.
        /// </summary>
        /// <param name="node">The node to render.</param>
        /// <param name="path">The path to save the rendered image to.</param>
        /// <param name="imageOptions">The image options to use during rendering. This can be null.</param>
        public static void RenderNode(Node node, string filePath, ImageSaveOptions imageOptions)
        {
            // Run some argument checks.
            if (node == null)
                throw new ArgumentException("Node cannot be null");

            // If no image options are supplied, create default options.
            if (imageOptions == null)
                imageOptions = new ImageSaveOptions(FileFormatUtil.ExtensionToSaveFormat(Path.GetExtension(filePath)));

            // Store the paper color to be used on the final image and change to transparent.
            // This will cause any content around the rendered node to be removed later on.
            Color savePaperColor = imageOptions.PaperColor;
            imageOptions.PaperColor = Color.Transparent;

            // There a bug which affects the cache of a cloned node. To avoid this we instead clone the entire document including all nodes,
            // find the matching node in the cloned document and render that instead.
            Document doc = (Document)node.Document.Clone(true);
            node = doc.GetChild(NodeType.Any, node.Document.GetChildNodes(NodeType.Any, true).IndexOf(node), true);

            // Create a temporary shape to store the target node in. This shape will be rendered to retrieve
            // the rendered content of the node.
            Shape shape = new Shape(doc, ShapeType.TextBox);
            Section parentSection = (Section)node.GetAncestor(NodeType.Section);

            // Assume that the node cannot be larger than the page in size.
            shape.Width = parentSection.PageSetup.PageWidth;
            shape.Height = parentSection.PageSetup.PageHeight;
            shape.FillColor = Color.Transparent; // We must make the shape and paper color transparent.

            // Don't draw a surronding line on the shape.
            shape.Stroked = false;

            // Move up through the DOM until we find node which is suitable to insert into a Shape (a node with a parent can contain paragraph, tables the same as a shape).
            // Each parent node is cloned on the way up so even a descendant node passed to this method can be rendered. 
            // Since we are working with the actual nodes of the document we need to clone the target node into the temporary shape.
            Node currentNode = node;
            while (!(currentNode.ParentNode is InlineStory || currentNode.ParentNode is Story || currentNode.ParentNode is ShapeBase))
            {
                CompositeNode parent = (CompositeNode)currentNode.ParentNode.Clone(false);
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
            shape.GetShapeRenderer().Save(stream, imageOptions);
            shape.Remove();

            // Load the image into a new bitmap.
            using (Bitmap renderedImage = new Bitmap(stream))
            {
                // Extract the actual content of the image by cropping transparent space around
                // the rendered shape.
                Rectangle cropRectangle = FindBoundingBoxAroundNode(renderedImage);

                Bitmap croppedImage = new Bitmap(cropRectangle.Width, cropRectangle.Height);
                croppedImage.SetResolution(imageOptions.Resolution, imageOptions.Resolution);

                // Create the final image with the proper background color.
                using (Graphics g = Graphics.FromImage(croppedImage))
                {
                    g.Clear(savePaperColor);
                    g.DrawImage(renderedImage, new Rectangle(0, 0, croppedImage.Width, croppedImage.Height), cropRectangle.X, cropRectangle.Y, cropRectangle.Width, cropRectangle.Height, GraphicsUnit.Pixel);
                    croppedImage.Save(filePath);
                }
            }
        }

        /// <summary>
        /// Finds the minimum bounding box around non-transparent pixels in a Bitmap.
        /// </summary>
        public static Rectangle FindBoundingBoxAroundNode(Bitmap originalBitmap)
        {
            Point min = new Point(int.MaxValue, int.MaxValue);
            Point max = new Point(int.MinValue, int.MinValue);

            for (int x = 0; x < originalBitmap.Width; ++x)
            {
                for (int y = 0; y < originalBitmap.Height; ++y)
                {
                    // Note that you can speed up this part of the algorithm by using LockBits and unsafe code instead of GetPixel.
                    Color pixelColor = originalBitmap.GetPixel(x, y);

                    // For each pixel that is not transparent calculate the bounding box around it.
                    if (pixelColor.ToArgb() != Color.Empty.ToArgb())
                    {
                        min.X = Math.Min(x, min.X);
                        min.Y = Math.Min(y, min.Y);
                        max.X = Math.Max(x, max.X);
                        max.Y = Math.Max(y, max.Y);
                    }
                }
            }

            // Add one pixel to the width and height to avoid clipping.
            return new Rectangle(min.X, min.Y, (max.X - min.X) + 1, (max.Y - min.Y) + 1);
        }
        //ExEnd
    }
}

