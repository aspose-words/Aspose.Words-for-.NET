using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;
using Aspose.Words.Drawing;
using System.Drawing;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Images
{
    class RemoveWatermark
    {
        // ExStart:RemoveWatermark
        public static void Run()
        {          
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithImages();
            string fileName = "RemoveWatermark.docx";
            Document doc = new Document(dataDir + fileName);
            RemoveWatermarkText(doc);
            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            doc.Save(dataDir);
        }

        private static void RemoveWatermarkText(Document doc)
        {
            foreach (HeaderFooter hf in doc.GetChildNodes(NodeType.HeaderFooter, true))
            {
                foreach (Shape shape in hf.GetChildNodes(NodeType.Shape, true))
                {
                    if (shape.Name.Contains("WaterMark"))
                    {
                        shape.Remove();
                    }
                }
            }
        }
    }
    // ExEnd:RemoveWatermark
}