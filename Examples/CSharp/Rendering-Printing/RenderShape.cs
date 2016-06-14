
using System;
using System.IO;
using System.Drawing;

using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using System.Drawing.Imaging;
using Aspose.Words.Tables;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class RenderShape
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

            // Load the documents which store the shapes we want to render.
            Document doc = new Document(dataDir + "TestFile RenderShape.doc");
            Document doc2 = new Document(dataDir + "TestFile RenderShape.docx");

            // Retrieve the target shape from the document. In our sample document this is the first shape.
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            // Test rendering of different types of nodes.
            RenderShapeToDisk(dataDir, shape);
            RenderShapeToStream(dataDir, shape);
            RenderShapeToGraphics(dataDir, shape);
            RenderCellToImage(dataDir, doc);
            RenderRowToImage(dataDir, doc);
            RenderParagraphToImage(dataDir, doc);
            FindShapeSizes(shape);
            RenderShapeImage(dataDir, shape);
        }
        public static void RenderShapeToDisk(string dataDir, Shape shape)
        {
            //ExStart:RenderShapeToDisk
            ShapeRenderer r = shape.GetShapeRenderer();

            // Define custom options which control how the image is rendered. Render the shape to the JPEG raster format.
            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Emf)
            {
                Scale = 1.5f
            };

            dataDir = dataDir + "TestFile.RenderToDisk_out_.emf";
            // Save the rendered image to disk.
            r.Save(dataDir, imageOptions);
            //ExEnd:RenderShapeToDisk
            Console.WriteLine("\nShape rendered to disk successfully.\nFile saved at " + dataDir);
        }
        public static void RenderShapeToStream(string dataDir, Shape shape)
        {
            //ExStart:RenderShapeToStream
            ShapeRenderer r = new ShapeRenderer(shape);

            // Define custom options which control how the image is rendered. Render the shape to the vector format EMF.
            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Output the image in gray scale
                ImageColorMode = ImageColorMode.Grayscale,

                // Reduce the brightness a bit (default is 0.5f).
                ImageBrightness = 0.45f
            };
            dataDir = dataDir + "TestFile.RenderToStream_out_.jpg";
            FileStream stream = new FileStream(dataDir, FileMode.Create);

            // Save the rendered image to the stream using different options.
            r.Save(stream, imageOptions);
            //ExEnd:RenderShapeToStream
            Console.WriteLine("\nShape rendered to stream successfully.\nFile saved at " + dataDir);
        }

        public static void RenderShapeToGraphics(string dataDir, Shape shape)
        {
            //ExStart:RenderShapeToGraphics
            ShapeRenderer r = shape.GetShapeRenderer();

            // Find the size that the shape will be rendered to at the specified scale and resolution.
            Size shapeSizeInPixels = r.GetSizeInPixels(1.0f, 96.0f);

            // Rotating the shape may result in clipping as the image canvas is too small. Find the longest side
            // and make sure that the graphics canvas is large enough to compensate for this.
            int maxSide = System.Math.Max(shapeSizeInPixels.Width, shapeSizeInPixels.Height);

            using (Bitmap image = new Bitmap((int)(maxSide * 1.25), (int)(maxSide * 1.25)))
            {
                // Rendering to a graphics object means we can specify settings and transformations to be applied to 
                // the shape that is rendered. In our case we will rotate the rendered shape.
                using (Graphics gr = Graphics.FromImage(image))
                {
                    // Clear the shape with the background color of the document.
                    gr.Clear(shape.Document.PageColor);
                    // Center the rotation using translation method below
                    gr.TranslateTransform((float)image.Width / 8, (float)image.Height / 2);
                    // Rotate the image by 45 degrees.
                    gr.RotateTransform(45);
                    // Undo the translation.
                    gr.TranslateTransform(-(float)image.Width / 8, -(float)image.Height / 2);

                    // Render the shape onto the graphics object.
                    r.RenderToSize(gr, 0, 0, shapeSizeInPixels.Width, shapeSizeInPixels.Height);
                }
                dataDir = dataDir + "TestFile.RenderToGraphics_out_.png";
                image.Save(dataDir, ImageFormat.Png);
                Console.WriteLine("\nShape rendered to graphics successfully.\nFile saved at " + dataDir);
            }
            //ExEnd:RenderShapeToGraphics
           
        }
        public static void RenderCellToImage(string dataDir, Document doc)
        {
            //ExStart:RenderCellToImage
            Cell cell = (Cell)doc.GetChild(NodeType.Cell, 2, true); // The third cell in the first table.
            dataDir = dataDir + "TestFile.RenderCell_out_.png";
            RenderNode(cell, dataDir, null);
            //ExEnd:RenderCellToImage
            Console.WriteLine("\nCell rendered to image successfully.\nFile saved at " + dataDir);
        }

        public static void RenderRowToImage(string dataDir, Document doc)
        {
            //ExStart:RenderRowToImage
            Row row = (Row)doc.GetChild(NodeType.Row, 0, true); // The first row in the first table.

            dataDir = dataDir + "TestFile.RenderRow_out_.png";
            RenderNode(row, dataDir, null);
            //ExEnd:RenderRowToImage
            Console.WriteLine("\nRow rendered to image successfully.\nFile saved at " + dataDir);
        }

        public static void RenderParagraphToImage(string dataDir, Document doc)
        {
            //ExStart:RenderParagraphToImage
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Paragraph paragraph = (Paragraph)shape.LastParagraph;

            // Save the node with a light pink background.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
            options.PaperColor = Color.LightPink;
            dataDir = dataDir + "TestFile.RenderParagraph_out_.png";
            RenderNode(paragraph, dataDir, options);
            //ExEnd:RenderParagraphToImage
            Console.WriteLine("\nParagraph rendered to image successfully.\nFile saved at " + dataDir);
        }
        public static void FindShapeSizes(Shape shape)
        {
            //ExStart:FindShapeSizes
            SizeF shapeSizeInDocument = shape.GetShapeRenderer().SizeInPoints;
            float width = shapeSizeInDocument.Width; // The width of the shape.
            float height = shapeSizeInDocument.Height; // The height of the shape.
            
            Size shapeRenderedSize = shape.GetShapeRenderer().GetSizeInPixels(1.0f, 96.0f);

            using (Bitmap image = new Bitmap(shapeRenderedSize.Width, shapeRenderedSize.Height))
            {
                using (Graphics g = Graphics.FromImage(image))
                {
                    // Render shape onto the graphics object using the RenderToScale or RenderToSize methods of ShapeRenderer class.
                }
            }
            //ExEnd:FindShapeSizes
        }
        public static void RenderShapeImage(string dataDir, Shape shape)
        {
            //ExStart:RenderShapeImage
            dataDir = dataDir + "TestFile.RenderShape_out_.jpg";
            // Save the Shape image to disk in JPEG format and using default options.
            shape.GetShapeRenderer().Save(dataDir, null);
            //ExEnd:RenderShapeImage
            Console.WriteLine("\nShape image rendered successfully.\nFile saved at " + dataDir);
        }
        /// <summary>
        /// Renders any node in a document to the path specified using the image save options.
        /// </summary>
        /// <param name="node">The node to render.</param>
        /// <param name="path">The path to save the rendered image to.</param>
        /// <param name="imageOptions">The image options to use during rendering. This can be null.</param>
        public static void RenderNode(Node node, string filePath, ImageSaveOptions imageOptions)
        {
            // This code is taken from public API samples of AW.
            // Previously to find opaque bounds of the shape the function
            // that checks every pixel of the rendered image was used.
            // For now opaque bounds is got using ShapeRenderer.GetOpaqueRectangleInPixels method.

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
            ShapeRenderer renderer = shape.GetShapeRenderer();
            renderer.Save(stream, imageOptions);
            shape.Remove();

            Rectangle crop = renderer.GetOpaqueBoundsInPixels(imageOptions.Scale, imageOptions.Resolution);

            // Load the image into a new bitmap.
            using (Bitmap renderedImage = new Bitmap(stream))
            {
                Bitmap croppedImage = new Bitmap(crop.Width, crop.Height);
                croppedImage.SetResolution(imageOptions.Resolution, imageOptions.Resolution);

                // Create the final image with the proper background color.
                using (Graphics g = Graphics.FromImage(croppedImage))
                {
                    g.Clear(savePaperColor);
                    g.DrawImage(renderedImage, new Rectangle(0, 0, croppedImage.Width, croppedImage.Height), crop.X, crop.Y, crop.Width, crop.Height, GraphicsUnit.Pixel);

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
                        min.X = System.Math.Min(x, min.X);
                        min.Y = System.Math.Min(y, min.Y);
                        max.X = System.Math.Max(x, max.X);
                        max.Y = System.Math.Max(y, max.Y);
                    }
                }
            }

            // Add one pixel to the width and height to avoid clipping.
            return new Rectangle(min.X, min.Y, (max.X - min.X) + 1, (max.Y - min.Y) + 1);
        }
    }
}
