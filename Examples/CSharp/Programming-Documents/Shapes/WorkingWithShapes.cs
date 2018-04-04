using Aspose.Words.Drawing;
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
            SetAspectRatioLocked();

        }

        public static void SetAspectRatioLocked()
        {
            // ExStart:SetAspectRatioLocked
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithShapes();

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
            watermark.IsLayoutInCell = false; // Display the shape outside of table cell if it will be placed into a cell.

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
    }
}
