using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Rendering;
using SkiaSharp;
using System.Drawing;
using System.IO;

namespace XamarinAndroid.Rendering_Printing
{
    class RenderShape
    {
        public static string Run()
        {
            string fileRelativePath = "/Data/Rendering-Printing/TestFile RenderShape.doc";
            string fileName = FileHelper.GetFileNameInAppData(fileRelativePath);

            fileName = FileHelper.GetFile(fileRelativePath);

            // Load the documents which store the shapes we want to render.
            Document doc = new Document(fileName);

            // Retrieve the target shape from the document. In our sample document this is the first shape.
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            // Test rendering of different types of nodes.
            return RenderShapeToGraphics(Path.GetDirectoryName(fileName), shape);
        }

		// ExStart:RenderShapeToGraphics
        public static string RenderShapeToGraphics(string dataDir, Shape shape)
        {
            ShapeRenderer r = shape.GetShapeRenderer();

            // Find the size that the shape will be rendered to at the specified scale and resolution.
            Size shapeSizeInPixels = r.GetSizeInPixels(1.0f, 96.0f);

            // Rotating the shape may result in clipping as the image canvas is too small. Find the longest side
            // And make sure that the graphics canvas is large enough to compensate for this.
            int maxSide = System.Math.Max(shapeSizeInPixels.Width, shapeSizeInPixels.Height);

            using (SkiaSharp.SKBitmap bitmap = new SkiaSharp.SKBitmap((int)(maxSide * 1.25), (int)(maxSide * 1.25)))
            {
                // Rendering to a graphics object means we can specify settings and transformations to be applied to 
                // The shape that is rendered. In our case we will rotate the rendered shape.
                using (SkiaSharp.SKCanvas gr = new SkiaSharp.SKCanvas(bitmap))
                {
                    // Clear the shape with the background color of the document.
                    gr.DrawColor(new SkiaSharp.SKColor(shape.Document.PageColor.R, shape.Document.PageColor.G, shape.Document.PageColor.B, shape.Document.PageColor.A));
                    // Center the rotation using translation method below
                    gr.Translate((float)bitmap.Width / 8, (float)bitmap.Height / 2);
                    // Rotate the image by 45 degrees.
                    gr.RotateDegrees(45);
                    // Undo the translation.
                    gr.Translate(-(float)bitmap.Width / 8, -(float)bitmap.Height / 2);

                    // Render the shape onto the graphics object.
                    r.RenderToSize(gr, 0, 0, shapeSizeInPixels.Width, shapeSizeInPixels.Height);
                }

                // Save output to file.
                using (System.IO.FileStream fs = System.IO.File.Create(dataDir + "/RenderToSize_Out.png"))
                {
                    SKData d = SKImage.FromBitmap(bitmap).Encode(SKEncodedImageFormat.Png, 100);
                    d.SaveTo(fs);
                }
            }

            return "\nShape rendered to graphics successfully.\nFile saved at " + dataDir;            
        }
		// ExEnd:RenderShapeToGraphics
    }
}