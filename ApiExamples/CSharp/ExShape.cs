﻿// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Drawing.Ole;
using Aspose.Words.Math;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;

namespace ApiExamples
{
    /// <summary>
    /// Examples using shapes in documents.
    /// </summary>
    [TestFixture]
    public class ExShape : ApiExampleBase
    {
        [Test]
        public void DeleteAllShapes()
        {
            Document doc = new Document(MyDir + "Shape.DeleteAllShapes.doc");

            //ExStart
            //ExFor:Shape
            //ExSummary:Shows how to delete all shapes from a document.
            // Here we get all shapes from the document node, but you can do this for any smaller
            // node too, for example delete shapes from a single section or a paragraph.
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            shapes.Clear();

            // There could also be group shapes, they have different node type, remove them all too.
            NodeCollection groupShapes = doc.GetChildNodes(NodeType.GroupShape, true);
            groupShapes.Clear();
            //ExEnd

            Assert.AreEqual(0, doc.GetChildNodes(NodeType.Shape, true).Count);
            Assert.AreEqual(0, doc.GetChildNodes(NodeType.GroupShape, true).Count);

            doc.Save(MyDir + @"\Artifacts\Shape.DeleteAllShapes.doc");
        }

        [Test]
        public void CheckShapeInline()
        {
            //ExStart
            //ExFor:ShapeBase.IsInline
            //ExSummary:Shows how to test if a shape in the document is inline or floating.
            Document doc = new Document(MyDir + "Shape.DeleteAllShapes.doc");

            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                if (shape.IsInline)
                    Console.WriteLine("Shape is inline.");
                else
                    Console.WriteLine("Shape is floating.");
            }
            //ExEnd

            // Verify that the first shape in the document is not inline.
            Assert.False(((Shape)doc.GetChild(NodeType.Shape, 0, true)).IsInline);
        }

        [Test]
        public void LineFlipOrientation()
        {
            //ExStart
            //ExFor:ShapeBase.Bounds
            //ExFor:ShapeBase.FlipOrientation
            //ExFor:FlipOrientation
            //ExSummary:Creates two line shapes. One line goes from top left to bottom right. Another line goes from bottom left to top right.
            Document doc = new Document();

            // The lines will cross the whole page.
            float pageWidth = (float)doc.FirstSection.PageSetup.PageWidth;
            float pageHeight = (float)doc.FirstSection.PageSetup.PageHeight;

            // This line goes from top left to bottom right by default. 
            Shape lineA = new Shape(doc, ShapeType.Line);
            lineA.Bounds = new RectangleF(0, 0, pageWidth, pageHeight);
            lineA.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            lineA.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            doc.FirstSection.Body.FirstParagraph.AppendChild(lineA);

            // This line goes from bottom left to top right because we flipped it. 
            Shape lineB = new Shape(doc, ShapeType.Line);
            lineB.Bounds = new RectangleF(0, 0, pageWidth, pageHeight);
            lineB.FlipOrientation = FlipOrientation.Horizontal;
            lineB.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            lineB.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            doc.FirstSection.Body.FirstParagraph.AppendChild(lineB);

            doc.Save(MyDir + @"\Artifacts\Shape.LineFlipOrientation.doc");
            //ExEnd
        }

        [Test]
        public void Fill()
        {
            //ExStart
            //ExFor:Shape.Fill
            //ExFor:Shape.FillColor
            //ExFor:Fill
            //ExFor:Fill.Opacity
            //ExSummary:Demonstrates how to create shapes with fill.
            DocumentBuilder builder = new DocumentBuilder();

            builder.Writeln();
            builder.Writeln();
            builder.Writeln();
            builder.Write("Some text under the shape.");

            // Create a red balloon, semitransparent.
            // The shape is floating and its coordinates are (0,0) by default, relative to the current paragraph.
            Shape shape = new Shape(builder.Document, ShapeType.Balloon);
            shape.FillColor = Color.Red;
            shape.Fill.Opacity = 0.3;
            shape.Width = 100;
            shape.Height = 100;
            shape.Top = -100;
            builder.InsertNode(shape);

            builder.Document.Save(MyDir + @"\Artifacts\Shape.Fill.doc");
            //ExEnd
        }

        [Test]
        public void GetShapeAltTextTitle()
        {
            //ExStart
            //ExFor:ShapeBase.Title
            //ExSummary:Shows how to get or set title of shape object.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create test shape.
            Shape shape = new Shape(doc, ShapeType.Cube);
            shape.Width = 431.5;
            shape.Height = 346.35;
            shape.Title = "Alt Text Title";

            builder.InsertNode(shape);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Console.WriteLine("Shape text: " + shape.Title);
            //ExEnd

            Assert.AreEqual("Alt Text Title", shape.Title);
        }

        [Test]
        public void ReplaceTextboxesWithImages()
        {
            //ExStart
            //ExFor:WrapSide
            //ExFor:ShapeBase.WrapSide
            //ExFor:NodeCollection
            //ExFor:CompositeNode.InsertAfter(Node, Node)
            //ExFor:NodeCollection.ToArray
            //ExSummary:Shows how to replace all textboxes with images.
            Document doc = new Document(MyDir + "Shape.ReplaceTextboxesWithImages.doc");

            // This gets a live collection of all shape nodes in the document.
            NodeCollection shapeCollection = doc.GetChildNodes(NodeType.Shape, true);

            // Since we will be adding/removing nodes, it is better to copy all collection
            // into a fixed size array, otherwise iterator will be invalidated.
            Node[] shapes = shapeCollection.ToArray();

            foreach (Shape shape in shapes)
            {
                // Filter out all shapes that we don't need.
                if (shape.ShapeType.Equals(ShapeType.TextBox))
                {
                    // Create a new shape that will replace the existing shape.
                    Shape image = new Shape(doc, ShapeType.Image);

                    // Load the image into the new shape.
                    image.ImageData.SetImage(ImageDir + "Hammer.wmf");

                    // Make new shape's position to match the old shape.
                    image.Left = shape.Left;
                    image.Top = shape.Top;
                    image.Width = shape.Width;
                    image.Height = shape.Height;
                    image.RelativeHorizontalPosition = shape.RelativeHorizontalPosition;
                    image.RelativeVerticalPosition = shape.RelativeVerticalPosition;
                    image.HorizontalAlignment = shape.HorizontalAlignment;
                    image.VerticalAlignment = shape.VerticalAlignment;
                    image.WrapType = shape.WrapType;
                    image.WrapSide = shape.WrapSide;

                    // Insert new shape after the old shape and remove the old shape.
                    shape.ParentNode.InsertAfter(image, shape);
                    shape.Remove();
                }
            }

            doc.Save(MyDir + @"\Artifacts\Shape.ReplaceTextboxesWithImages.doc");
            //ExEnd
        }

        [Test]
        public void CreateTextBox()
        {
            //ExStart
            //ExFor:Shape.#ctor(DocumentBase, ShapeType)
            //ExFor:ShapeBase.ZOrder
            //ExFor:Story.FirstParagraph
            //ExFor:Shape.FirstParagraph
            //ExFor:ShapeBase.WrapType
            //ExSummary:Creates a textbox with some text and different formatting options in a new document.
            // Create a blank document.
            Document doc = new Document();

            // Create a new shape of type TextBox
            Shape textBox = new Shape(doc, ShapeType.TextBox);

            // Set some settings of the textbox itself.
            // Set the wrap of the textbox to inline
            textBox.WrapType = WrapType.None;
            // Set the horizontal and vertical alignment of the text inside the shape.
            textBox.HorizontalAlignment = HorizontalAlignment.Center;
            textBox.VerticalAlignment = VerticalAlignment.Top;

            // Set the textbox height and width.
            textBox.Height = 50;
            textBox.Width = 200;

            // Set the textbox in front of other shapes with a lower ZOrder
            textBox.ZOrder = 2;

            // Let's create a new paragraph for the textbox manually and align it in the center. Make sure we add the new nodes to the textbox as well.
            textBox.AppendChild(new Paragraph(doc));
            Paragraph para = textBox.FirstParagraph;
            para.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            // Add some text to the paragraph.
            Run run = new Run(doc);
            run.Text = "Content in textbox";
            para.AppendChild(run);

            // Append the textbox to the first paragraph in the body.
            doc.FirstSection.Body.FirstParagraph.AppendChild(textBox);

            // Save the output
            doc.Save(MyDir + @"\Artifacts\Shape.CreateTextBox.doc");
            //ExEnd
        }

        [Test]
        public void GetActiveXControlProperties()
        {
            //ExStart
            //ExFor:OleControl
            //ExFor:Forms2OleControl.Caption
            //ExFor:Forms2OleControl.Value
            //ExFor:Forms2OleControl.Enabled
            //ExFor:Forms2OleControl.Type
            //ExFor:Forms2OleControl.ChildNodes
            //ExSummary: Shows how to get ActiveX control and properties from the document.
            Document doc = new Document(MyDir + "Shape.ActiveXObject.docx");

            //Get ActiveX control from the document 
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            OleControl oleControl = shape.OleFormat.OleControl;

            //Get ActiveX control properties
            if (oleControl.IsForms2OleControl)
            {
                Forms2OleControl checkBox = (Forms2OleControl)oleControl;
                Assert.AreEqual("Первый", checkBox.Caption);
                Assert.AreEqual("0", checkBox.Value);
                Assert.AreEqual(true, checkBox.Enabled);
                Assert.AreEqual(Forms2OleControlType.CheckBox, checkBox.Type);
                Assert.AreEqual(null, checkBox.ChildNodes);
            }
            //ExEnd
        }

        [Test]
        public void SuggestedFileName()
        {
            //ExStart
            //ExFor:OleFormat.SuggestedFileName
            //ExSummary:Shows how to get suggested file name from the object.
            Document doc = new Document(MyDir + "Shape.SuggestedFileName.rtf");

            // Gets the file name suggested for the current embedded object if you want to save it into a file
            Shape oleShape = (Shape)doc.FirstSection.Body.GetChild(NodeType.Shape, 0, true);
            String suggestedFileName = oleShape.OleFormat.SuggestedFileName;
            //ExEnd

            Assert.AreEqual("CSV.csv", suggestedFileName);
        }

        [Test]
        public void ObjectDidNotHaveSuggestedFileName()
        {
            Document doc = new Document(MyDir + "Shape.ActiveXObject.docx");

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Assert.IsEmpty(shape.OleFormat.SuggestedFileName);
        }

        [Test]
        public void GetOpaqueBoundsInPixels()
        {
            Document doc = new Document(MyDir + "Shape.TextBox.doc");

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Jpeg);

            MemoryStream stream = new MemoryStream();
            ShapeRenderer renderer = shape.GetShapeRenderer();
            renderer.Save(stream, imageOptions);

            shape.Remove();

            // Check that the opaque bounds and bounds have default values
            Assert.AreEqual(250, renderer.GetOpaqueBoundsInPixels(imageOptions.Scale, imageOptions.VerticalResolution).Width);
            Assert.AreEqual(52, renderer.GetOpaqueBoundsInPixels(imageOptions.Scale, imageOptions.HorizontalResolution).Height);

            Assert.AreEqual(250, renderer.GetBoundsInPixels(imageOptions.Scale, imageOptions.VerticalResolution).Width);
            Assert.AreEqual(52, renderer.GetBoundsInPixels(imageOptions.Scale, imageOptions.HorizontalResolution).Height);

            Assert.AreEqual(250, renderer.GetOpaqueBoundsInPixels(imageOptions.Scale, imageOptions.HorizontalResolution).Width);
            Assert.AreEqual(52, renderer.GetOpaqueBoundsInPixels(imageOptions.Scale, imageOptions.HorizontalResolution).Height);

            Assert.AreEqual(250, renderer.GetBoundsInPixels(imageOptions.Scale, imageOptions.VerticalResolution).Width);
            Assert.AreEqual(52, renderer.GetBoundsInPixels(imageOptions.Scale, imageOptions.VerticalResolution).Height);

            Assert.AreEqual((float)187.849991, renderer.OpaqueBoundsInPoints.Width);
            Assert.AreEqual((float)39.25, renderer.OpaqueBoundsInPoints.Height);
        }

        [Test]
        public void ResolutionDefaultValues()
        {
            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Jpeg);

            Assert.AreEqual(96, imageOptions.HorizontalResolution);
            Assert.AreEqual(96, imageOptions.VerticalResolution);
        }

        //For assert result of the test you need to open "Shape.OfficeMath.svg" and check that OfficeMath node is there
        [Test]
        public void SaveShapeObjectAsImage()
        {
            //ExStart
            //ExFor:OfficeMath.GetMathRenderer
            //ExFor:NodeRendererBase.Save(String, ImageSaveOptions)
            //ExSummary:Shows how to convert specific object into image
            Document doc = new Document(MyDir + "Shape.OfficeMath.docx");

            //Get OfficeMath node from the document and render this as image (you can also do the same with the Shape node)
            OfficeMath math = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
            math.GetMathRenderer().Save(MyDir + @"\Artifacts\Shape.OfficeMath.svg", new ImageSaveOptions(SaveFormat.Svg));
            //ExEnd
        }

        [Test]
        public void OfficeMathDisplayException()
        {
            Document doc = new Document(MyDir + "Shape.OfficeMath.docx");

            OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
            officeMath.DisplayType = OfficeMathDisplayType.Display;

            Assert.That(() => officeMath.Justification = OfficeMathJustification.Inline, Throws.TypeOf<ArgumentException>());
        }

        [Test]
        public void OfficeMathDefaultValue()
        {
            Document doc = new Document(MyDir + "Shape.OfficeMath.docx");

            OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);

            Assert.AreEqual(OfficeMathDisplayType.Display, officeMath.DisplayType);
            Assert.AreEqual(OfficeMathJustification.Center, officeMath.Justification);
        }

        [Test]
        public void OfficeMathDisplayGold()
        {
            //ExStart
            //ExFor:OfficeMath.DisplayType
            //ExFor:OfficeMath.Justification
            //ExSummary:Shows how to set office math display formatting.
            Document doc = new Document(MyDir + "Shape.OfficeMath.docx");

            OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
            officeMath.DisplayType = OfficeMathDisplayType.Display;
            officeMath.Justification = OfficeMathJustification.Left;

            doc.Save(MyDir + @"Artifacts\Shape.OfficeMath.docx");
            //ExEnd
            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"Artifacts\Shape.OfficeMath.docx", MyDir + @"\Golds\Shape.OfficeMath Gold.docx"));
        }

        [Test]
        public void CannotBeSetDisplayWithInlineJustification()
        {
            Document doc = new Document(MyDir + "Shape.OfficeMath.docx");

            OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
            officeMath.DisplayType = OfficeMathDisplayType.Display;

            Assert.Throws<ArgumentException>(() => officeMath.Justification = OfficeMathJustification.Inline);
        }

        [Test]
        public void CannotBeSetInlineDisplayWithJustification()
        {
            Document doc = new Document(MyDir + "Shape.OfficeMath.docx");

            OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);
            officeMath.DisplayType = OfficeMathDisplayType.Inline;

            Assert.Throws<ArgumentException>(() => officeMath.Justification = OfficeMathJustification.Center);
        }

        [Test]
        public void OfficeMathDisplayNestedObjects()
        {
            Document doc = new Document(MyDir + "Shape.NestedOfficeMath.docx");

            OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, 0, true);

            //Always inline
            Assert.AreEqual(OfficeMathDisplayType.Inline, officeMath.DisplayType);
            Assert.AreEqual(OfficeMathJustification.Inline, officeMath.Justification);
        }

        [Test]
        [TestCase(0, MathObjectType.OMathPara)]
        [TestCase(1, MathObjectType.OMath)]
        [TestCase(2, MathObjectType.Supercript)]
        [TestCase(3, MathObjectType.Argument)]
        [TestCase(4, MathObjectType.SuperscriptPart)]
        public void WorkWithMathObjectType(int index, MathObjectType objectType)
        {
            Document doc = new Document(MyDir + "Shape.OfficeMath.docx");

            OfficeMath officeMath = (OfficeMath)doc.GetChild(NodeType.OfficeMath, index, true);
            Assert.AreEqual(objectType, officeMath.MathObjectType);
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void AspectRatioLocked(bool isLocked)
        {
            //ExStart
            //ExFor:ShapeBase.AspectRatioLocked
            //ExSummary:Shows how to set "AspectRatioLocked" for the shape object
            Document doc = new Document(MyDir + "Shape.ActiveXObject.docx");

            // Get shape object from the document and set AspectRatioLocked(it is possible to get/set AspectRatioLocked for child shapes (mimic MS Word behavior), 
            // but AspectRatioLocked has effect only for top level shapes!)
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            shape.AspectRatioLocked = isLocked;
            //ExEnd

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Assert.AreEqual(isLocked, shape.AspectRatioLocked);
        }

        [Test]
        public void AspectRatioLockedDefaultValue()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The best place for the watermark image is in the header or footer so it is shown on every page.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            Image image = Image.FromFile(ImageDir + "Watermark.png");

            // Insert a floating picture.
            Shape shape = builder.InsertImage(image);
            shape.WrapType = WrapType.None;
            shape.BehindText = true;

            shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;

            // Calculate image left and top position so it appears in the centre of the page.
            shape.Left = (builder.PageSetup.PageWidth - shape.Width) / 2;
            shape.Top = (builder.PageSetup.PageHeight - shape.Height) / 2;

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Assert.AreEqual(true, shape.AspectRatioLocked);
        }

        [Test]
        public void MarkupLunguageByDefault()
        {
            //ExStart
            //ExFor:ShapeBase.MarkupLanguage
            //ExSummary:Shows how get markup language for shape object in document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape image = builder.InsertImage(ImageDir + "dotnet-logo.png");

            // Loop through all single shapes inside document.
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                Assert.AreEqual(ShapeMarkupLanguage.Dml, shape.MarkupLanguage);

                Console.WriteLine("Shape: " + shape.MarkupLanguage);
                Console.WriteLine("ShapeSize: " + shape.SizeInPoints);
            }
            //ExEnd
        }

        [Test]
        [TestCase(MsWordVersion.Word2000, ShapeMarkupLanguage.Vml)]
        [TestCase(MsWordVersion.Word2002, ShapeMarkupLanguage.Vml)]
        [TestCase(MsWordVersion.Word2003, ShapeMarkupLanguage.Vml)]
        [TestCase(MsWordVersion.Word2007, ShapeMarkupLanguage.Vml)]
        [TestCase(MsWordVersion.Word2010, ShapeMarkupLanguage.Dml)]
        [TestCase(MsWordVersion.Word2013, ShapeMarkupLanguage.Dml)]
        [TestCase(MsWordVersion.Word2016, ShapeMarkupLanguage.Dml)]
        public void MarkupLunguageForDifferentMsWordVersions(MsWordVersion msWordVersion, ShapeMarkupLanguage shapeMarkupLanguage)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            doc.CompatibilityOptions.OptimizeFor(msWordVersion);

            Shape image = builder.InsertImage(ImageDir + "dotnet-logo.png");

            // Loop through all single shapes inside document.
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                Assert.AreEqual(shapeMarkupLanguage, shape.MarkupLanguage);
            }
        }

        [Test]
        public void ChangeStrokeProperties()
        {
            //ExStart
            //ExFor:Stroke
            //ExSummary:Shows how change stroke properties
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a new shape of type Rectangle
            Shape rectangle = new Shape(doc, ShapeType.Rectangle);

            //Change stroke properties
            Stroke stroke = rectangle.Stroke;
            stroke.On = true;
            stroke.Weight = 5;
            stroke.Color = Color.Red;
            stroke.DashStyle = DashStyle.ShortDashDotDot;
            stroke.JoinStyle = JoinStyle.Miter;
            stroke.EndCap = EndCap.Square;
            stroke.LineStyle = ShapeLineStyle.Triple;

            //Insert shape object
            builder.InsertNode(rectangle);
            //ExEnd

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            rectangle = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Stroke strokeAfter = rectangle.Stroke;

            Assert.AreEqual(true, strokeAfter.On);
            Assert.AreEqual(5, strokeAfter.Weight);
            Assert.AreEqual(Color.Red.ToArgb(), strokeAfter.Color.ToArgb());
            Assert.AreEqual(DashStyle.ShortDashDotDot, strokeAfter.DashStyle);
            Assert.AreEqual(JoinStyle.Miter, strokeAfter.JoinStyle);
            Assert.AreEqual(EndCap.Square, strokeAfter.EndCap);
            Assert.AreEqual(ShapeLineStyle.Triple, strokeAfter.LineStyle);
        }

        [Test]
        [Description("WORDSNET-16067")]
        public void InsertOleObjectAsHtmlFile()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertOleObject("http://www.aspose.com", "htmlfile", true, false, null);

            doc.Save(MyDir + @"\Artifacts\Document.InsertedOleObject.docx");
        }

        [Test]
        [Description("WORDSNET-16085")]
        public void InsertOlePackage()
        {
            //ExStart
            //ExFor:OlePackage
            //ExFor:OleFormat.OlePackage
            //ExFor:OlePackage.FileName
            //ExFor:OlePackage.DisplayName
            //ExSummary:Shows how insert ole object as ole package and set it file name and display name.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            byte[] zipFileBytes = File.ReadAllBytes(DatabaseDir + "cat001.zip");

            using (MemoryStream stream = new MemoryStream(zipFileBytes))
            {
                Shape shape = builder.InsertOleObject(stream, "Package", true, null);

                OlePackage setOlePackage = shape.OleFormat.OlePackage;
                setOlePackage.FileName = "Cat FileName.zip";
                setOlePackage.DisplayName = "Cat DisplayName.zip";

                doc.Save(MyDir + @"\Artifacts\Shape.InsertOlePackage.docx");
            }
            //ExEnd

            doc = new Document(MyDir + @"\Artifacts\Shape.InsertOlePackage.docx");

            Shape getShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            OlePackage getOlePackage = getShape.OleFormat.OlePackage;

            Assert.AreEqual("Cat FileName.zip", getOlePackage.FileName);
            Assert.AreEqual("Cat DisplayName.zip", getOlePackage.DisplayName);
        }

        [Test]
        public void GetAccessToOlePackage()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape oleObject = builder.InsertOleObject(MyDir + "Document.Spreadsheet.xlsx", false, false, null);
            Shape oleObjectAsOlePackage = builder.InsertOleObject(MyDir + "Document.Spreadsheet.xlsx", "Excel.Sheet", false, false, null);

            Assert.AreEqual(null, oleObject.OleFormat.OlePackage);
            Assert.AreEqual(typeof(OlePackage), oleObjectAsOlePackage.OleFormat.OlePackage.GetType());
        }

        [Test]
        public void NumberFormat()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add chart with default data.
            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);
            Chart chart = shape.Chart;
            chart.Title.Text = "Data Labels With Different Number Format";

            // Delete default generated series.
            chart.Series.Clear();

            // Add new series
            ChartSeries series0 = chart.Series.Add("AW Series 0", new[] { "AW0", "AW1", "AW2" }, new[] { 2.5, 1.5, 3.5 });

            // Add DataLabel to the first point of the first series.
            ChartDataLabel chartDataLabel0 = series0.DataLabels.Add(0);
            chartDataLabel0.ShowValue = true;

            // Set currency format code.
            chartDataLabel0.NumberFormat.FormatCode = "\"$\"#,##0.00";

            ChartDataLabel chartDataLabel1 = series0.DataLabels.Add(1);
            chartDataLabel1.ShowValue = true;

            // Set date format code.
            chartDataLabel1.NumberFormat.FormatCode = "d/mm/yyyy";

            ChartDataLabel chartDataLabel2 = series0.DataLabels.Add(2);
            chartDataLabel2.ShowValue = true;

            // Set percentage format code.
            chartDataLabel2.NumberFormat.FormatCode = "0.00%";

            // Or you can set format code to be linked to a source cell,
            // in this case NumberFormat will be reset to general and inherited from a source cell.
            chartDataLabel2.NumberFormat.IsLinkedToSource = true;

            doc.Save(MyDir + @"\Artifacts\DocumentBuilder.NumberFormat.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\DocumentBuilder.NumberFormat.docx", MyDir + @"\Golds\DocumentBuilder.NumberFormat Gold.docx"));
        }

        [Test]
        public void DataArraysWrongSize()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add chart with default data.
            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);
            Chart chart = shape.Chart;

            ChartSeriesCollection seriesColl = chart.Series;
            seriesColl.Clear();

            // Create category names array, second category will be null.
            string[] categories = { "Cat1", null, "Cat3", "Cat4", "Cat5", null };

            // Adding new series with empty (double.NaN) values.
            seriesColl.Add("AW Series 1", categories, new double[] { 1, 2, double.NaN, 4, 5, 6 });
            seriesColl.Add("AW Series 2", categories, new double[] { 2, 3, double.NaN, 5, 6, 7 });

            Assert.That(() => seriesColl.Add("AW Series 3", categories, new[] { double.NaN, 4, 5, double.NaN, double.NaN }), Throws.TypeOf<ArgumentException>());
            Assert.That(() => seriesColl.Add("AW Series 4", categories, new[] { double.NaN, double.NaN, double.NaN, double.NaN, double.NaN }), Throws.TypeOf<ArgumentException>());
        }

        [Test]
        public void EmptyValuesInChartData()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add chart with default data.
            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);
            Chart chart = shape.Chart;

            ChartSeriesCollection seriesColl = chart.Series;
            seriesColl.Clear();

            // Create category names array, second category will be null.
            string[] categories = { "Cat1", null, "Cat3", "Cat4", "Cat5", null };

            // Adding new series with empty (double.NaN) values.
            seriesColl.Add("AW Series 1", categories, new[] { 1, 2, double.NaN, 4, 5, 6 });
            seriesColl.Add("AW Series 2", categories, new[] { 2, 3, double.NaN, 5, 6, 7 });
            seriesColl.Add("AW Series 3", categories, new[] { double.NaN, 4, 5, double.NaN, 7, 8 });
            seriesColl.Add("AW Series 4", categories, new[] { double.NaN, double.NaN, double.NaN, double.NaN, double.NaN, 9 });

            doc.Save(MyDir + @"\Artifacts\EmptyValuesInChartData.docx");
        }

        [Test]
        public void ChartDefaultValues()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert chart.
            builder.InsertChart(ChartType.Column3D, 432, 252);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Shape shapeNode = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Chart chart = shapeNode.Chart;
            
            // Assert X axis
            Assert.AreEqual(ChartAxisType.Category, chart.AxisX.Type);
            Assert.AreEqual(AxisCategoryType.Automatic, chart.AxisX.CategoryType);
            Assert.AreEqual(AxisCrosses.Automatic, chart.AxisX.Crosses);
            Assert.AreEqual(false, chart.AxisX.ReverseOrder);
            Assert.AreEqual(AxisTickMark.None, chart.AxisX.MajorTickMark);
            Assert.AreEqual(AxisTickMark.None, chart.AxisX.MinorTickMark);
            Assert.AreEqual(AxisTickLabelPosition.NextToAxis, chart.AxisX.TickLabelPosition);
            Assert.AreEqual(1, chart.AxisX.MajorUnit);
            Assert.AreEqual(true, chart.AxisX.MajorUnitIsAuto);
            Assert.AreEqual(AxisTimeUnit.Automatic, chart.AxisX.MajorUnitScale);
            Assert.AreEqual(0.5, chart.AxisX.MinorUnit);
            Assert.AreEqual(true, chart.AxisX.MinorUnitIsAuto);
            Assert.AreEqual(AxisTimeUnit.Automatic, chart.AxisX.MinorUnitScale);
            Assert.AreEqual(AxisTimeUnit.Automatic, chart.AxisX.BaseTimeUnit);
            Assert.AreEqual("General", chart.AxisX.NumberFormat.FormatCode);
            Assert.AreEqual(100, chart.AxisX.TickLabelOffset);
            Assert.AreEqual(AxisBuiltInUnit.None, chart.AxisX.DisplayUnit.Unit);
            Assert.AreEqual(true, chart.AxisX.AxisBetweenCategories);
            Assert.AreEqual(AxisScaleType.Linear, chart.AxisX.Scaling.Type);
            Assert.AreEqual(1, chart.AxisX.TickLabelSpacing);
            Assert.AreEqual(true, chart.AxisX.TickLabelSpacingIsAuto);
            Assert.AreEqual(1, chart.AxisX.TickMarkSpacing);

            // Assert Y axis
            Assert.AreEqual(ChartAxisType.Value, chart.AxisY.Type);
            Assert.AreEqual(AxisCategoryType.Category, chart.AxisY.CategoryType);
            Assert.AreEqual(AxisCrosses.Automatic, chart.AxisY.Crosses);
            Assert.AreEqual(false, chart.AxisY.ReverseOrder);
            Assert.AreEqual(AxisTickMark.None, chart.AxisY.MajorTickMark);
            Assert.AreEqual(AxisTickMark.None, chart.AxisY.MinorTickMark);
            Assert.AreEqual(AxisTickLabelPosition.NextToAxis, chart.AxisY.TickLabelPosition);
            Assert.AreEqual(1, chart.AxisY.MajorUnit);
            Assert.AreEqual(true, chart.AxisY.MajorUnitIsAuto);
            Assert.AreEqual(AxisTimeUnit.Automatic, chart.AxisY.MajorUnitScale);
            Assert.AreEqual(0.5, chart.AxisY.MinorUnit);
            Assert.AreEqual(true, chart.AxisY.MinorUnitIsAuto);
            Assert.AreEqual(AxisTimeUnit.Automatic, chart.AxisY.MinorUnitScale);
            Assert.AreEqual(AxisTimeUnit.Automatic, chart.AxisY.BaseTimeUnit);
            Assert.AreEqual("General", chart.AxisY.NumberFormat.FormatCode);
            Assert.AreEqual(100, chart.AxisY.TickLabelOffset);
            Assert.AreEqual(AxisBuiltInUnit.None, chart.AxisY.DisplayUnit.Unit);
            Assert.AreEqual(true, chart.AxisY.AxisBetweenCategories);
            Assert.AreEqual(AxisScaleType.Linear, chart.AxisY.Scaling.Type);
            Assert.AreEqual(1, chart.AxisY.TickLabelSpacing);
            Assert.AreEqual(true, chart.AxisY.TickLabelSpacingIsAuto);
            Assert.AreEqual(1, chart.AxisY.TickMarkSpacing);

            // Assert Z axis
            Assert.AreEqual(ChartAxisType.Series, chart.AxisZ.Type);
            Assert.AreEqual(AxisCategoryType.Category, chart.AxisZ.CategoryType);
            Assert.AreEqual(AxisCrosses.Automatic, chart.AxisZ.Crosses);
            Assert.AreEqual(false, chart.AxisZ.ReverseOrder);
            Assert.AreEqual(AxisTickMark.None, chart.AxisZ.MajorTickMark);
            Assert.AreEqual(AxisTickMark.None, chart.AxisZ.MinorTickMark);
            Assert.AreEqual(AxisTickLabelPosition.NextToAxis, chart.AxisZ.TickLabelPosition);
            Assert.AreEqual(1, chart.AxisZ.MajorUnit);
            Assert.AreEqual(true, chart.AxisZ.MajorUnitIsAuto);
            Assert.AreEqual(AxisTimeUnit.Automatic, chart.AxisZ.MajorUnitScale);
            Assert.AreEqual(0.5, chart.AxisZ.MinorUnit);
            Assert.AreEqual(true, chart.AxisZ.MinorUnitIsAuto);
            Assert.AreEqual(AxisTimeUnit.Automatic, chart.AxisZ.MinorUnitScale);
            Assert.AreEqual(AxisTimeUnit.Automatic, chart.AxisZ.BaseTimeUnit);
            Assert.AreEqual(string.Empty, chart.AxisZ.NumberFormat.FormatCode);
            Assert.AreEqual(100, chart.AxisZ.TickLabelOffset);
            Assert.AreEqual(AxisBuiltInUnit.None, chart.AxisZ.DisplayUnit.Unit);
            Assert.AreEqual(true, chart.AxisZ.AxisBetweenCategories);
            Assert.AreEqual(AxisScaleType.Linear, chart.AxisZ.Scaling.Type);
            Assert.AreEqual(1, chart.AxisZ.TickLabelSpacing);
            Assert.AreEqual(true, chart.AxisZ.TickLabelSpacingIsAuto);
            Assert.AreEqual(1, chart.AxisZ.TickMarkSpacing);
        }

        [Test]
        public void InsertChartUsingAxisProperties()
        {
            //ExStart
            //ExFor:ChartAxis
            //ExFor:ChartAxis.CategoryType
            //ExFor:ChartAxis.Crosses
            //ExFor:ChartAxis.ReverseOrder
            //ExFor:ChartAxis.MajorTickMark
            //ExFor:ChartAxis.MinorTickMark
            //ExFor:ChartAxis.MajorUnit
            //ExFor:ChartAxis.MinorUnit
            //ExFor:ChartAxis.TickLabelOffset
            //ExFor:ChartAxis.TickLabelPosition
            //ExFor:ChartAxis.TickLabelSpacingIsAuto
            //ExFor:ChartAxis.TickMarkSpacing
            //ExSummary:Shows how to insert chart using the axis options for detailed configuration.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert chart.
            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = shape.Chart;

            // Clear demo data.
            chart.Series.Clear();

            chart.Series.Add("Aspose Test Series", 
                new string[] { "Word", "PDF", "Excel", "GoogleDocs", "Note"}, 
                new double[] { 640, 320, 280, 120, 150 });

            // Get chart axises
            ChartAxis xAxis = chart.AxisX;
            ChartAxis yAxis = chart.AxisY;

            // Set X-axis options
            xAxis.CategoryType = AxisCategoryType.Category;
            xAxis.Crosses = AxisCrosses.Minimum;
            xAxis.ReverseOrder = false;
            xAxis.MajorTickMark = AxisTickMark.Inside;
            xAxis.MinorTickMark = AxisTickMark.Cross;
            xAxis.MajorUnit = 10;
            xAxis.MinorUnit = 15;
            xAxis.TickLabelOffset = 50;
            xAxis.TickLabelPosition = AxisTickLabelPosition.Low;
            xAxis.TickLabelSpacingIsAuto = false;
            xAxis.TickMarkSpacing = 1;

            // Set Y-axis options
            yAxis.CategoryType = AxisCategoryType.Automatic;
            yAxis.Crosses = AxisCrosses.Maximum;
            yAxis.ReverseOrder = true;
            yAxis.MajorTickMark = AxisTickMark.Inside;
            yAxis.MinorTickMark = AxisTickMark.Cross;
            yAxis.MajorUnit = 100;
            yAxis.MinorUnit = 20;
            yAxis.TickLabelPosition = AxisTickLabelPosition.NextToAxis;
            //ExEnd
            
            doc.Save(MyDir + @"\Artifacts\Shape.InsertChartUsingAxisProperties Out.docx");
            doc.Save(MyDir + @"\Artifacts\Shape.InsertChartUsingAxisProperties Out.pdf");
        }

        [Test]
        public void InsertChartWithDateTimeValues()
        {
            //ExStart
            //ExFor:ChartAxis.Scaling
            //ExFor:AxisScaling.Minimum
            //ExFor:AxisScaling.Maximum
            //ExSummary: Shows how to insert chart with date/time values
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert chart.
            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);
            Chart chart = shape.Chart;

            // Clear demo data.
            chart.Series.Clear();

            // Fill data.
            chart.Series.Add("Aspose Test Series",
                new DateTime[] { new DateTime(2017, 11, 06), new DateTime(2017, 11, 09), new DateTime(2017, 11, 15),
                    new DateTime(2017, 11, 21), new DateTime(2017, 11, 25), new DateTime(2017, 11, 29) },
                new double[] { 1.2, 0.3, 2.1, 2.9, 4.2, 5.3 });

            // Set X axis bounds.
            ChartAxis xAxis = chart.AxisX;
            xAxis.Scaling.Minimum = new DateTime(2017, 11, 05).ToOADate();
            xAxis.Scaling.Maximum = new DateTime(2017, 12, 03).ToOADate();

            // Set major units to a week and minor units to a day.
            xAxis.MajorUnit = 7;
            xAxis.MinorUnit = 1;
            xAxis.MajorTickMark = AxisTickMark.Cross;
            xAxis.MinorTickMark = AxisTickMark.Outside;
            //ExEnd

            doc.Save(MyDir + @"\Artifacts\Shape.InsertChartWithDateTimeValues Out.docx");
            doc.Save(MyDir + @"\Artifacts\Shape.InsertChartWithDateTimeValues Out.pdf");
        }

        [Test]
        public void SetNumberFormatToChartAxis()
        {
            //ExStart
            //ExFor:ChartAxis.NumberFormat
            //ExFor:NumberFormat.FormatCode
            //ExSummary:Shows how to set formatting for chart values.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert chart.
            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = shape.Chart;

            // Clear demo data.
            chart.Series.Clear();

            chart.Series.Add("Aspose Test Series",
                new string[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
                new double[] { 1900000, 850000, 2100000, 600000, 1500000 });

            // Set number format.
            chart.AxisY.NumberFormat.FormatCode = "#,##0";
            //ExEnd

            doc.Save(MyDir + @"\Artifacts\Shape.SetNumberFormatToChartAxis Out.docx");
            doc.Save(MyDir + @"\Artifacts\Shape.SetNumberFormatToChartAxis Out.pdf");
        }

        // Note: Tests below used for verification conversion docx to pdf and the correct display.
        // For now, the results check manually.
        [Test]
        [TestCase(ChartType.Column)]
        [TestCase(ChartType.Line)]
        [TestCase(ChartType.Pie)]
        [TestCase(ChartType.Bar)]
        [TestCase(ChartType.Area)]
        public void TestDisplayChartsWithConversion(ChartType chartType)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert chart.
            Shape shape = builder.InsertChart(chartType, 432, 252);
            Chart chart = shape.Chart;

            // Clear demo data.
            chart.Series.Clear();

            chart.Series.Add("Aspose Test Series",
                new string[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
                new double[] { 1900000, 850000, 2100000, 600000, 1500000 });

            doc.Save(MyDir + @"\Artifacts\Shape.TestDisplayChartsWithConversion Out.docx");
            doc.Save(MyDir + @"\Artifacts\Shape.TestDisplayChartsWithConversion Out.pdf");
        }
        
        [Test]
        public void Surface3DChart()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert chart.
            Shape shape = builder.InsertChart(ChartType.Surface3D, 432, 252);
            Chart chart = shape.Chart;

            // Clear demo data.
            chart.Series.Clear();

            chart.Series.Add("Aspose Test Series 1",
                new string[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
                new double[] { 1900000, 850000, 2100000, 600000, 1500000 });

            chart.Series.Add("Aspose Test Series 2",
                new string[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
                new double[] { 900000, 50000, 1100000, 400000, 2500000 });

            chart.Series.Add("Aspose Test Series 3",
                new string[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
                new double[] { 500000, 820000, 1500000, 400000, 100000 });

            doc.Save(MyDir + @"\Artifacts\SurfaceChart Out.docx");
            doc.Save(MyDir + @"\Artifacts\SurfaceChart Out.pdf");
        }
        
        [Test]
        public void BubbleChart()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert chart.
            Shape shape = builder.InsertChart(ChartType.Bubble, 432, 252);
            Chart chart = shape.Chart;

            // Clear demo data.
            chart.Series.Clear();

            chart.Series.Add("Aspose Test Series",
                new double[] { 2900000, 350000, 1100000, 400000, 400000 },
                new double[] { 1900000, 850000, 2100000, 600000, 1500000 },
                new double[] { 900000, 450000, 2500000, 800000, 500000 });

            doc.Save(MyDir + @"\Artifacts\BubbleChart Out.docx");
            doc.Save(MyDir + @"\Artifacts\BubbleChart Out.pdf");
        }
    }
}
