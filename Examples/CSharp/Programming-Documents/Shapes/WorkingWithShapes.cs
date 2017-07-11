using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Shapes
{
    class WorkingWithShapes
    {
        public static void Run()
        {
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
    }
}
