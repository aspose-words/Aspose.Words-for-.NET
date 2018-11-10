using System;
using System.Drawing;
using System.Net;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;
using Shape = Aspose.Words.Drawing.Shape;

namespace ApiExamples
{
    [TestFixture]
    public class ExDrawing : ApiExampleBase
    {
#if !(NETSTANDARD2_0 || __MOBILE__)
        [Test]
        public void DrawingVariousShapes()
        {
            //ExStart
            //ExFor:Drawing.ArrowLength
            //ExFor:Drawing.ArrowType
            //ExFor:Drawing.ArrowWidth
            //ExFor:Drawing.DashStyle
            //ExFor:Drawing.EndCap
            //ExFor:Drawing.Fill.Color
            //ExFor:Drawing.Fill.ImageBytes
            //ExFor:Drawing.Fill.On
            //ExFor:Drawing.JoinStyle
            //ExFor:Stroke.Color
            //ExFor:Stroke.StartArrowLength
            //ExFor:Stroke.StartArrowType
            //ExFor:Stroke.StartArrowWidth
            //ExFor:Stroke.DashStyle
            //ExFor:Stroke.EndArrowType
            //ExFor:Stroke.EndCap
            //ExSummary:Shows to create a variety of shapes
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Draw a dotted horizontal red line with an arrow on the left end and a diamond on the other
            Shape arrow = new Shape(doc, ShapeType.Line);
            arrow.Width = 200;
            arrow.Stroke.Color = Color.Red;
            arrow.Stroke.StartArrowType = ArrowType.Arrow;
            arrow.Stroke.StartArrowLength = ArrowLength.Long;
            arrow.Stroke.StartArrowWidth = ArrowWidth.Wide;
            arrow.Stroke.EndArrowType = ArrowType.Diamond;
            arrow.Stroke.DashStyle = DashStyle.Dash;

            Assert.AreEqual(JoinStyle.Miter, arrow.Stroke.JoinStyle);

            builder.InsertNode(arrow);

            // Draw a thick black diagonal line with rounded ends
            Shape line = new Shape(doc, ShapeType.Line);
            line.Top = 40;
            line.Width = 200;
            line.Height = 20;
            line.StrokeWeight = 5.0;
            line.Stroke.EndCap = EndCap.Round;

            builder.InsertNode(line);

            // Draw an arrow with a green fill
            Shape filledInArrow = new Shape(doc, ShapeType.Arrow);
            filledInArrow.Width = 200;
            filledInArrow.Height = 40;
            filledInArrow.Top = 100;
            filledInArrow.Fill.Color = Color.Green;
            filledInArrow.Fill.On = true;

            builder.InsertNode(filledInArrow);

            // Draw an arrow filled in with the Aspose logo and flip its orientation
            Shape filledInArrowImg = new Shape(doc, ShapeType.Arrow);
            filledInArrowImg.Width = 200;
            filledInArrowImg.Height = 40;
            filledInArrowImg.Top = 160;
            filledInArrowImg.FlipOrientation = FlipOrientation.Both;

            using (WebClient webClient = new WebClient())
            {
                byte[] imageBytes = webClient.DownloadData("http://www.aspose.com/images/aspose-logo.gif");

                using (System.IO.MemoryStream stream = new System.IO.MemoryStream(imageBytes))
                {
                    Image image = Image.FromStream(stream);
                    // When we flipped the orientation of our arrow, the image content was flipped too
                    // If we want it to be displayed the right side up, we have to reverse the arrow flip on the image
                    image.RotateFlip(RotateFlipType.RotateNoneFlipXY);

                    filledInArrowImg.ImageData.SetImage(image);
                    builder.InsertNode(filledInArrowImg);

                    filledInArrowImg.Stroke.JoinStyle = JoinStyle.Round;
                }
            }

            doc.Save(MyDir + @"\Artifacts\Drawing.VariousShapes.docx");
            //ExEnd
        }
#endif

        //ExStart
        //ExFor:DocumentVisitor.VisitShapeEnd(Shape)
        //ExFor:DocumentVisitor.VisitShapeStart(Shape)
        //ExFor:DocumentVisitor.VisitGroupShapeEnd(GroupShape)
        //ExFor:DocumentVisitor.VisitGroupShapeStart(GroupShape)
        //ExFor:Drawing.GroupShape
        //ExFor:Drawing.GroupShape.#ctor(DocumentBase)
        //ExFor:Drawing.GroupShape.#ctor(DocumentBase,Drawing.ShapeMarkupLanguage)
        //ExFor:Drawing.GroupShape.Accept(DocumentVisitor)
        //ExSummary:Shows how to create a group of shapes, and let it accept a visitor
        [Test] //ExSkip
        public void GroupOfShapes()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            GroupShape group = new GroupShape(doc);

            Assert.AreEqual(0, group.ChildNodes.Count);

            Shape balloon = new Shape(doc, ShapeType.Balloon);
            balloon.Width = 200;
            balloon.Height = 200;
            balloon.Stroke.Color = Color.Red;

            Shape cube = new Shape(doc, ShapeType.Cube);
            cube.Width = 100;
            cube.Height = 100;
            cube.Stroke.Color = Color.Blue;

            group.AppendChild(balloon);
            group.AppendChild(cube);

            builder.InsertNode(group);

            ShapeInfoPrinter printer = new ShapeInfoPrinter();

            group.Accept(printer);

            Console.WriteLine(printer.GetText());
        }

        /// <summary>
        /// Visitor that prints shape group contents information to the console.
        /// </summary>
        public class ShapeInfoPrinter : DocumentVisitor
        {
            public ShapeInfoPrinter()
            {
                mBuilder = new StringBuilder();
            }

            public string GetText()
            {
                return mBuilder.ToString();
            }

            public override VisitorAction VisitGroupShapeStart(GroupShape groupShape)
            {
                mBuilder.AppendLine("Shape group started:");
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitGroupShapeEnd(GroupShape groupShape)
            {
                mBuilder.AppendLine("End of shape group");
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitShapeStart(Shape shape)
            {
                mBuilder.AppendLine("\tShape - " + shape.ShapeType + ":");
                mBuilder.AppendLine("\t\tWidth: " + shape.Width);
                mBuilder.AppendLine("\t\tHeight: " + shape.Height);
                mBuilder.AppendLine("\t\tStroke color: " + shape.Stroke.Color);
                mBuilder.AppendLine("\t\tFill color: " + shape.Fill.Color);
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitShapeEnd(Shape shape)
            {
                mBuilder.AppendLine("\tEnd of shape");
                return VisitorAction.Continue;
            }

            private readonly StringBuilder mBuilder;
        }
        //ExEnd

#if !(NETSTANDARD2_0 || __MOBILE__)
        [Test]
        public void TypeOfImage()
        {
            //ExStart
            //ExFor:Drawing.ImageType
            //ExSummary:Shows how to add an image to a shape and check its type
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            using (WebClient webClient = new WebClient())
            {
                byte[] imageBytes = webClient.DownloadData("http://www.aspose.com/images/aspose-logo.gif");

                using (System.IO.MemoryStream stream = new System.IO.MemoryStream(imageBytes))
                {
                    Image image = Image.FromStream(stream);

                    // The image started off as an animated .gif but it gets converted to a .png since there cannot be animated images in documents
                    Shape imgShape = builder.InsertImage(image);
                    Assert.AreEqual(ImageType.Png, imgShape.ImageData.ImageType);
                }
            }

            //ExEnd
        }
#endif

        [Test]
        public void TextBoxTextLayout()
        {
            //ExStart
            //ExFor:Drawing.LayoutFlow
            //ExSummary:Shows how to add text to a textbox and change its orientation
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape textbox = new Shape(doc, ShapeType.TextBox);

            textbox.Width = 100;
            textbox.Height = 100;

            textbox.AppendChild(new Paragraph(doc));

            builder.InsertNode(textbox);

            builder.MoveTo(textbox.FirstParagraph);

            builder.Write("This text is flipped 90 degrees to the left.");

            textbox.TextBox.LayoutFlow = LayoutFlow.BottomToTop;
            doc.Save(MyDir + @"\Artifacts\Drawing.TextBox.docx");
            //ExEnd
        }
    }
}