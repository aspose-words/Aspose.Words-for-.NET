// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Rendering;
using NUnit.Framework;
#if NET462 || JAVA
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing.Text;
#elif NETCOREAPP2_1 || __MOBILE__
using SkiaSharp;
#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExRendering : ApiExampleBase
    {
#if NET462 || JAVA
        //ExStart
        //ExFor:NodeRendererBase.RenderToScale(Graphics, Single, Single, Single)
        //ExFor:NodeRendererBase.RenderToSize(Graphics, Single, Single, Single, Single)
        //ExFor:ShapeRenderer
        //ExFor:ShapeRenderer.#ctor(ShapeBase)
        //ExSummary:Shows how to render a shape with a Graphics object and display it using a Windows Form.
        [Test, Category("IgnoreOnJenkins"), Category("SkipMono")] //ExSkip
        public void RenderShapesOnForm()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            ShapeForm shapeForm = new ShapeForm(new Size(1017, 840));

            // Below are two ways to use the "ShapeRenderer" class to render a shape to a Graphics object.
            // 1 -  Create a shape with a chart, and render it to a specific scale.
            Chart chart = builder.InsertChart(ChartType.Pie, 500, 400).Chart;
            chart.Series.Clear();
            chart.Series.Add("Desktop Browser Market Share (Oct. 2020)",
                new[] { "Google Chrome", "Apple Safari", "Mozilla Firefox", "Microsoft Edge", "Other" },
                new[] { 70.33, 8.87, 7.69, 5.83, 7.28 });

            Shape chartShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            shapeForm.AddShapeToRenderToScale(chartShape, 0, 0, 1.5f);

            // 2 -  Create a shape group, and render it to a specific size.
            GroupShape group = new GroupShape(doc);
            group.Bounds = new RectangleF(0, 0, 100, 100);
            group.CoordSize = new Size(500, 500);

            Shape subShape = new Shape(doc, ShapeType.Rectangle);
            subShape.Width = 500;
            subShape.Height = 500;
            subShape.Left = 0;
            subShape.Top = 0;
            subShape.FillColor = Color.RoyalBlue;
            group.AppendChild(subShape);

            subShape = new Shape(doc, ShapeType.Image);
            subShape.Width = 450;
            subShape.Height = 450;
            subShape.Left = 25;
            subShape.Top = 25;
            subShape.ImageData.SetImage(ImageDir + "Logo.jpg");
            group.AppendChild(subShape);

            builder.InsertNode(group);

            GroupShape groupShape = (GroupShape)doc.GetChild(NodeType.GroupShape, 0, true);
            shapeForm.AddShapeToRenderToSize(groupShape, 880, 680, 100, 100);

            shapeForm.ShowDialog();
        }

        /// <summary>
        /// Renders and displays a list of shapes.
        /// </summary>
        private class ShapeForm : Form
        {
            public ShapeForm(Size size)
            {
                Timer timer = new Timer(); //ExSKip
                timer.Interval = 10000; //ExSKip
                timer.Tick += TimerTick; //ExSKip
                timer.Start(); //ExSKip
                Size = size;
                mShapesToRender = new List<KeyValuePair<ShapeBase, float[]>>();
            }

            public void AddShapeToRenderToScale(ShapeBase shape, float x, float y, float scale)
            {
                mShapesToRender.Add(new KeyValuePair<ShapeBase, float[]>(shape, new[] {x, y, scale}));
            }

            public void AddShapeToRenderToSize(ShapeBase shape, float x, float y, float width, float height)
            {
                mShapesToRender.Add(new KeyValuePair<ShapeBase, float[]>(shape, new[] {x, y, width, height}));
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                foreach (KeyValuePair<ShapeBase, float[]> renderingArgs in mShapesToRender)
                    if (renderingArgs.Value.Length == 3)
                        RenderShapeToScale(renderingArgs.Key, renderingArgs.Value[0], renderingArgs.Value[1],
                            renderingArgs.Value[2]);
                    else if (renderingArgs.Value.Length == 4)
                        RenderShapeToSize(renderingArgs.Key, renderingArgs.Value[0], renderingArgs.Value[1],
                            renderingArgs.Value[2], renderingArgs.Value[3]);
            }

            private void RenderShapeToScale(ShapeBase shape, float x, float y, float scale)
            {
                ShapeRenderer renderer = new ShapeRenderer(shape);
                using (Graphics formGraphics = CreateGraphics())
                {
                    renderer.RenderToScale(formGraphics, x, y, scale);
                }
            }

            private void RenderShapeToSize(ShapeBase shape, float x, float y, float width, float height)
            {
                ShapeRenderer renderer = new ShapeRenderer(shape);
                using (Graphics formGraphics = CreateGraphics())
                {
                    renderer.RenderToSize(formGraphics, x, y, width, height);
                }
            }

            private void TimerTick(object sender, EventArgs e) => Close(); //ExSkip
            private readonly List<KeyValuePair<ShapeBase, float[]>> mShapesToRender;
        }
        //ExEnd

        [Test]
        public void RenderToSize()
        {
            //ExStart
            //ExFor:Document.RenderToSize
            //ExSummary:Shows how to render a document to a bitmap at a specified location and size.
            Document doc = new Document(MyDir + "Rendering.docx");
            
            using (Bitmap bmp = new Bitmap(700, 700))
            {
                using (Graphics gr = Graphics.FromImage(bmp))
                {
                    gr.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;

                    // Set the "PageUnit" property to "GraphicsUnit.Inch" to use inches as the
                    // measurement unit for any transformations and dimensions that we will define.
                    gr.PageUnit = GraphicsUnit.Inch;

                    // Offset the output 0.5" from the edge.
                    gr.TranslateTransform(0.5f, 0.5f);

                    // Rotate the output by 10 degrees.
                    gr.RotateTransform(10);

                    // Draw a 3"x3" rectangle.
                    gr.DrawRectangle(new Pen(Color.Black, 3f / 72f), 0f, 0f, 3f, 3f);
                    
                    // Draw the first page of our document with the same dimensions and transformation as the rectangle.
                    // The rectangle will frame the first page.
                    float returnedScale = doc.RenderToSize(0, gr, 0f, 0f, 3f, 3f);

                    // This is the scaling factor that the RenderToSize method applied to the first page to fit the specified size.
                    Assert.AreEqual(0.2566f, returnedScale, 0.0001f);

                    // Set the "PageUnit" property to "GraphicsUnit.Millimeter" to use millimeters as the
                    // measurement unit for any transformations and dimensions that we will define.
                    gr.PageUnit = GraphicsUnit.Millimeter;

                    // Reset the transformations that we used from the previous rendering.
                    gr.ResetTransform();

                    // Apply another set of transformations. 
                    gr.TranslateTransform(10, 10);
                    gr.ScaleTransform(0.5f, 0.5f);
                    gr.PageScale = 2f;

                    // Create another rectangle and use it to frame another page from the document.
                    gr.DrawRectangle(new Pen(Color.Black, 1), 90, 10, 50, 100);
                    doc.RenderToSize(1, gr, 90, 10, 50, 100);

                    bmp.Save(ArtifactsDir + "Rendering.RenderToSize.png");
                }
            }
            //ExEnd
        }

        [Test]
        public void Thumbnails()
        {
            //ExStart
            //ExFor:Document.RenderToScale
            //ExSummary:Shows how to the individual pages of a document to graphics to create one image with thumbnails of all pages.
            Document doc = new Document(MyDir + "Rendering.docx");

            // Calculate the number of rows and columns that we will fill with thumbnails.
            const int thumbColumns = 2;
            int thumbRows = Math.DivRem(doc.PageCount, thumbColumns, out int remainder);

            if (remainder > 0)
                thumbRows++;

            // Scale the thumbnails relative to the size of the first page.
            const float scale = 0.25f;
            Size thumbSize = doc.GetPageInfo(0).GetSizeInPixels(scale, 96);

            // Calculate the size of the image that will contain all the thumbnails.
            int imgWidth = thumbSize.Width * thumbColumns;
            int imgHeight = thumbSize.Height * thumbRows;
            
            using (Bitmap img = new Bitmap(imgWidth, imgHeight))
            {
                using (Graphics gr = Graphics.FromImage(img))
                {
                    gr.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;

                    // Fill the background, which is transparent by default, in white.
                    gr.FillRectangle(new SolidBrush(Color.White), 0, 0, imgWidth, imgHeight);

                    for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
                    {
                        int rowIdx = Math.DivRem(pageIndex, thumbColumns, out int columnIdx);

                        // Specify where we want the thumbnail to appear.
                        float thumbLeft = columnIdx * thumbSize.Width;
                        float thumbTop = rowIdx * thumbSize.Height;

                        // Render a page as a thumbnail, and then frame it in a rectangle of the same size.
                        SizeF size = doc.RenderToScale(pageIndex, gr, thumbLeft, thumbTop, scale);
                        gr.DrawRectangle(Pens.Black, thumbLeft, thumbTop, size.Width, size.Height);
                    }

                    img.Save(ArtifactsDir + "Rendering.Thumbnails.png");
                }
            }
            //ExEnd
        }
#elif NETCOREAPP2_1 || __MOBILE__
        [Test]
        public void RenderToSizeNetStandard2()
        {
            //ExStart
            //ExFor:Document.RenderToSize
            //ExSummary:Shows how to render the document as a bitmap at a specified location and size (.NetStandard 2.0).
            Document doc = new Document(MyDir + "Rendering.docx");
            
            using (SKBitmap bitmap = new SKBitmap(700, 700))
            {
                using (SKCanvas canvas = new SKCanvas(bitmap))
                {
                    // Apply a scaling factor of 70% to the page that we will render using this canvas.
                    canvas.Scale(70);

                    // Offset the page 0.5" from the top and left edges of the page.
                    canvas.Translate(0.5f, 0.5f);

                    // Rotate the rendered page by 10 degrees.
                    canvas.RotateDegrees(10);

                    // Create and draw a rectangle.
                    SKRect rect = new SKRect(0f, 0f, 3f, 3f);
                    canvas.DrawRect(rect, new SKPaint
                    {
                        Color = SKColors.Black,
                        Style = SKPaintStyle.Stroke,
                        StrokeWidth = 3f / 72f
                    });

                    // Render the first page of the document to the same size as the above rectangle.
                    // The rectangle will frame this page.
                    float returnedScale = doc.RenderToSize(0, canvas, 0f, 0f, 3f, 3f);

                    Console.WriteLine("The image was rendered at {0:P0} zoom.", returnedScale);

                    // Reset the matrix, and then apply a new set of scaling and translations.
                    canvas.ResetMatrix();
                    canvas.Scale(5);
                    canvas.Translate(10, 10);

                    // Create another rectangle.
                    rect = new SKRect(0, 0, 50, 100);
                    rect.Offset(90, 10);
                    canvas.DrawRect(rect, new SKPaint
                    {
                        Color = SKColors.Black,
                        Style = SKPaintStyle.Stroke,
                        StrokeWidth = 1
                    });

                    // Render the first page within the newly created rectangle once again.
                    doc.RenderToSize(0, canvas, 90, 10, 50, 100);

                    using (SKFileWStream fs = new SKFileWStream(ArtifactsDir + "Rendering.RenderToSizeNetStandard2.png"))
                    {
                        bitmap.PeekPixels().Encode(fs, SKEncodedImageFormat.Png, 100);
                    }
                }
            }            
            //ExEnd
        }

        [Test]
        public void CreateThumbnailsNetStandard2()
        {
            //ExStart
            //ExFor:Document.RenderToScale
            //ExSummary:Renders individual pages to graphics to create one image with thumbnails of all pages (.NetStandard 2.0).
            Document doc = new Document(MyDir + "Rendering.docx");

            // Calculate the number of rows and columns that we will fill with thumbnails.
            const int thumbnailColumnsNum = 2;
            int thumbRows = Math.DivRem(doc.PageCount, thumbnailColumnsNum, out int remainder);

            if (remainder > 0)
                thumbRows++;

            // Scale the thumbnails relative to the size of the first page. 
            const float scale = 0.25f;
            Size thumbSize = doc.GetPageInfo(0).GetSizeInPixels(scale, 96);

            // Calculate the size of the image that will contain all the thumbnails.
            int imgWidth = thumbSize.Width * thumbnailColumnsNum;
            int imgHeight = thumbSize.Height * thumbRows;

            using (SKBitmap bitmap = new SKBitmap(imgWidth, imgHeight))
            {
                using (SKCanvas canvas = new SKCanvas(bitmap))
                {
                    // Fill the background, which is transparent by default, in white.
                    canvas.Clear(SKColors.White);

                    for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
                    {
                        int rowIdx = Math.DivRem(pageIndex, thumbnailColumnsNum, out int columnIdx);

                        // Specify where we want the thumbnail to appear.
                        float thumbLeft = columnIdx * thumbSize.Width;
                        float thumbTop = rowIdx * thumbSize.Height;

                        SizeF size = doc.RenderToScale(pageIndex, canvas, thumbLeft, thumbTop, scale);

                        // Render a page as a thumbnail, and then frame it in a rectangle of the same size.
                        SKRect rect = new SKRect(0, 0, size.Width, size.Height);
                        rect.Offset(thumbLeft, thumbTop);
                        canvas.DrawRect(rect, new SKPaint
                        {
                            Color = SKColors.Black,
                            Style = SKPaintStyle.Stroke
                        });
                    }

                    using (SKFileWStream fs = new SKFileWStream(ArtifactsDir + "Rendering.CreateThumbnailsNetStandard2.png"))
                    {
                        bitmap.PeekPixels().Encode(fs, SKEncodedImageFormat.Png, 100);
                    }
                }
            }            
            //ExEnd
        }
#endif
    }
}