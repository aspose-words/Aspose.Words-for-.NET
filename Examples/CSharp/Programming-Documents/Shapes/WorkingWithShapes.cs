using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Shapes
{
    public class WorkingWithShapes
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithShapes();
            SetShapeLayoutInCell(dataDir);
            SetAspectRatioLocked(dataDir);
            InsertShapeUsingDocumentBuilder(dataDir);
            AddCornersSnipped(dataDir);
            GetActualShapeBoundsPoints(dataDir);
            SpecifyVerticalAnchor(dataDir);
            DetectSmartArtShape(dataDir);
        }

        public static void InsertShapeUsingDocumentBuilder(string dataDir)
        {
            // ExStart:InsertShapeUsingDocumentBuilder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            //Free-floating shape insertion.
            Shape shape = builder.InsertShape(ShapeType.TextBox, RelativeHorizontalPosition.Page, 100, RelativeVerticalPosition.Page, 100, 50, 50, WrapType.None);
            shape.Rotation = 30.0;

            builder.Writeln();

            //Inline shape insertion.
            shape = builder.InsertShape(ShapeType.TextBox, 50, 50);
            shape.Rotation = 30.0;

            OoxmlSaveOptions so = new OoxmlSaveOptions(SaveFormat.Docx);
            // "Strict" or "Transitional" compliance allows to save shape as DML.
            so.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;

            dataDir = dataDir + "Shape_InsertShapeUsingDocumentBuilder_out.docx";

            // Save the document to disk.
            doc.Save(dataDir, so);
            // ExEnd:InsertShapeUsingDocumentBuilder
            Console.WriteLine("\nInsert Shape successfully using DocumentBuilder.\nFile saved at " + dataDir);
        }

        public static void SetAspectRatioLocked(string dataDir)
        {
            // ExStart:SetAspectRatioLocked
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            var shape = builder.InsertImage(dataDir + "Test.png");
            shape.AspectRatioLocked = false;

            dataDir = dataDir + "Shape_AspectRatioLocked_out.doc";

            // Save the document to disk.
            doc.Save(dataDir);
            // ExEnd:SetAspectRatioLocked
            Console.WriteLine("\nShape's AspectRatioLocked property is set successfully.\nFile saved at " + dataDir);
        }

        public static void SetShapeLayoutInCell(string dataDir)
        {
            // ExStart:SetShapeLayoutInCell

            Document doc = new Document(dataDir + @"LayoutInCell.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape watermark = new Shape(doc, ShapeType.TextPlainText);
            watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            watermark.IsLayoutInCell = true; // Display the shape outside of table cell if it will be placed into a cell.

            watermark.Width = 300;
            watermark.Height = 70;
            watermark.HorizontalAlignment = HorizontalAlignment.Center;
            watermark.VerticalAlignment = VerticalAlignment.Center;

            watermark.Rotation = -40;
            watermark.Fill.Color = Color.Gray;
            watermark.StrokeColor = Color.Gray;

            watermark.TextPath.Text = "watermarkText";
            watermark.TextPath.FontFamily = "Arial";

            watermark.Name = string.Format("WaterMark_{0}", Guid.NewGuid());
            watermark.WrapType = WrapType.None;

            Run run = doc.GetChildNodes(NodeType.Run, true)[doc.GetChildNodes(NodeType.Run, true).Count - 1] as Run;

            builder.MoveTo(run);
            builder.InsertNode(watermark);
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2010);

            dataDir = dataDir + "Shape_IsLayoutInCell_out.docx";

            // Save the document to disk.
            doc.Save(dataDir);
            // ExEnd:SetShapeLayoutInCell
            Console.WriteLine("\nShape's IsLayoutInCell property is set successfully.\nFile saved at " + dataDir);
        }

        public static void AddCornersSnipped(string dataDir)
        {
            // ExStart:AddCornersSnipped
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            Shape shape = builder.InsertShape(ShapeType.TopCornersSnipped, 50, 50);

            OoxmlSaveOptions so = new OoxmlSaveOptions(SaveFormat.Docx);
            so.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;
            dataDir = dataDir + "AddCornersSnipped_out.docx";

            //Save the document to disk.
            doc.Save(dataDir, so);
            // ExEnd:AddCornersSnipped
            Console.WriteLine("\nCorner Snip shape is created successfully.\nFile saved at " + dataDir);
        }

        public static void GetActualShapeBoundsPoints(string dataDir)
        {
            // ExStart:GetActualShapeBoundsPoints
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            var shape = builder.InsertImage(dataDir + "Test.png");
            shape.AspectRatioLocked = false;

            Console.Write("\nGets the actual bounds of the shape in points.");
            Console.WriteLine(shape.GetShapeRenderer().BoundsInPoints);
            // ExEnd:GetActualShapeBoundsPoints
        }

        public static void SpecifyVerticalAnchor(string dataDir)
        {
            // ExStart:SpecifyVerticalAnchor
            Document doc = new Document(dataDir + @"VerticalAnchor.docx");
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            Shape textBoxShape = shapes[0] as Shape;
            if (textBoxShape != null)
            {
                textBoxShape.TextBox.VerticalAnchor = TextBoxAnchor.Bottom;
            }
            doc.Save(dataDir + "VerticalAnchor_out.docx");
            // ExEnd:SpecifyVerticalAnchor
        }

        public static void DetectSmartArtShape(string dataDir)
        {
            // ExStart:DetectSmartArtShape
            Document doc = new Document(dataDir + "input.docx");

            int count = 0;
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                if (shape.HasSmartArt)
                    count++;
            }

            Console.WriteLine("The document has {0} shapes with SmartArt.", count);
            // ExEnd:DetectSmartArtShape
        }
    }
}
