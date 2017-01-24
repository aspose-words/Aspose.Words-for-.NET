
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
using Aspose.Words.Drawing.Ole;
namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class ReadActiveXControlProperties
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

            // Load the documents which store the shapes we want to render.           
            Document doc = new Document(dataDir + "ActiveXControl.docx");

            string properties = "";
            // Retrieve shapes from the document.         
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                OleControl oleControl = shape.OleFormat.OleControl;               
                if (oleControl.IsForms2OleControl)
                {
                    Forms2OleControl checkBox = (Forms2OleControl)oleControl;
                    properties = properties + "\nCaption: " + checkBox.Caption;
                    properties = properties + "\nValue: " + checkBox.Value;
                    properties = properties + "\nEnabled: " + checkBox.Enabled;
                    properties = properties + "\nType: " + checkBox.Type;
                    if (checkBox.ChildNodes != null)
                    {
                        properties = properties + "\nChildNodes: " + checkBox.ChildNodes;
                    }

                    properties = properties + "\n";
                }
            }
            properties = properties + "\nTotal ActiveX Controls found: " + doc.GetChildNodes(NodeType.Shape, true).Count.ToString();
            Console.WriteLine("\n" + properties);
        }      
    }
}
