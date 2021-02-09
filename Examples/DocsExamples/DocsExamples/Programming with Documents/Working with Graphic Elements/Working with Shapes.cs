using System;
using System.Drawing;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Working_with_Graphic_Elements
{
    internal class WorkingWithShapes : DocsExamplesBase
    {
        [Test]
        public void AddGroupShape()
        {
            //ExStart:AddGroupShape
            Document doc = new Document();
            doc.EnsureMinimum();
            
            GroupShape groupShape = new GroupShape(doc);
            Shape accentBorderShape = new Shape(doc, ShapeType.AccentBorderCallout1) { Width = 100, Height = 100 };
            groupShape.AppendChild(accentBorderShape);

            Shape actionButtonShape = new Shape(doc, ShapeType.ActionButtonBeginning)
            {
                Left = 100, Width = 100, Height = 200
            };
            groupShape.AppendChild(actionButtonShape);

            groupShape.Width = 200;
            groupShape.Height = 200;
            groupShape.CoordSize = new Size(200, 200);

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertNode(groupShape);

            doc.Save(ArtifactsDir + "WorkingWithShapes.AddGroupShape.docx");
            //ExEnd:AddGroupShape
        }

        [Test]
        public void InsertShape()
        {
            //ExStart:InsertShape
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertShape(ShapeType.TextBox, RelativeHorizontalPosition.Page, 100,
                RelativeVerticalPosition.Page, 100, 50, 50, WrapType.None);
            shape.Rotation = 30.0;

            builder.Writeln();

            shape = builder.InsertShape(ShapeType.TextBox, 50, 50);
            shape.Rotation = 30.0;

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                Compliance = OoxmlCompliance.Iso29500_2008_Transitional
            };

            doc.Save(ArtifactsDir + "WorkingWithShapes.InsertShape.docx", saveOptions);
            //ExEnd:InsertShape
        }

        [Test]
        public void AspectRatioLocked()
        {
            //ExStart:AspectRatioLocked
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertImage(ImagesDir + "Transparent background logo.png");
            shape.AspectRatioLocked = false;

            doc.Save(ArtifactsDir + "WorkingWithShapes.AspectRatioLocked.docx");
            //ExEnd:AspectRatioLocked
        }

        [Test]
        public void LayoutInCell()
        {
            //ExStart:LayoutInCell
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartTable();
            builder.RowFormat.Height = 100;
            builder.RowFormat.HeightRule = HeightRule.Exactly;

            for (int i = 0; i < 31; i++)
            {
                if (i != 0 && i % 7 == 0) builder.EndRow();
                builder.InsertCell();
                builder.Write("Cell contents");
            }

            builder.EndTable();

            Shape watermark = new Shape(doc, ShapeType.TextPlainText)
            {
                RelativeHorizontalPosition = RelativeHorizontalPosition.Page,
                RelativeVerticalPosition = RelativeVerticalPosition.Page,
                IsLayoutInCell = true, // Display the shape outside of the table cell if it will be placed into a cell.
                Width = 300,
                Height = 70,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Rotation = -40
            };

            watermark.Fill.Color = Color.Gray;
            watermark.StrokeColor = Color.Gray;

            watermark.TextPath.Text = "watermarkText";
            watermark.TextPath.FontFamily = "Arial";

            watermark.Name = $"WaterMark_{Guid.NewGuid()}";
            watermark.WrapType = WrapType.None;

            Run run = doc.GetChildNodes(NodeType.Run, true)[doc.GetChildNodes(NodeType.Run, true).Count - 1] as Run;

            builder.MoveTo(run);
            builder.InsertNode(watermark);
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

            doc.Save(ArtifactsDir + "WorkingWithShapes.LayoutInCell.docx");
            //ExEnd:LayoutInCell
        }

        [Test]
        public void AddCornersSnipped()
        {
            //ExStart:AddCornersSnipped
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertShape(ShapeType.TopCornersSnipped, 50, 50);

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
            {
                Compliance = OoxmlCompliance.Iso29500_2008_Transitional
            };

            doc.Save(ArtifactsDir + "WorkingWithShapes.AddCornersSnipped.docx", saveOptions);
            //ExEnd:AddCornersSnipped
        }

        [Test]
        public void GetActualShapeBoundsPoints()
        {
            //ExStart:GetActualShapeBoundsPoints
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertImage(ImagesDir + "Transparent background logo.png");
            shape.AspectRatioLocked = false;

            Console.Write("\nGets the actual bounds of the shape in points: ");
            Console.WriteLine(shape.GetShapeRenderer().BoundsInPoints);
            //ExEnd:GetActualShapeBoundsPoints
        }

        [Test]
        public void VerticalAnchor()
        {
            //ExStart:VerticalAnchor
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape textBox = builder.InsertShape(ShapeType.TextBox, 200, 200);
            textBox.TextBox.VerticalAnchor = TextBoxAnchor.Bottom;
            
            builder.MoveTo(textBox.FirstParagraph);
            builder.Write("Textbox contents");

            doc.Save(ArtifactsDir + "WorkingWithShapes.VerticalAnchor.docx");
            //ExEnd:VerticalAnchor
        }

        [Test]
        public void DetectSmartArtShape()
        {
            //ExStart:DetectSmartArtShape
            Document doc = new Document(MyDir + "SmartArt.docx");

            int count = doc.GetChildNodes(NodeType.Shape, true).Cast<Shape>().Count(shape => shape.HasSmartArt);

            Console.WriteLine("The document has {0} shapes with SmartArt.", count);
            //ExEnd:DetectSmartArtShape
        }

        [Test]
        public void UpdateSmartArtDrawing()
        {
            Document doc = new Document(MyDir + "SmartArt.docx");

            //ExStart:UpdateSmartArtDrawing
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
                if (shape.HasSmartArt)
                    shape.UpdateSmartArtDrawing();
            //ExEnd:UpdateSmartArtDrawing
        }
    }
}