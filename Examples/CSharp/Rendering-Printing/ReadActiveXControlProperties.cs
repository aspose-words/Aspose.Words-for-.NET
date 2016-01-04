//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
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
namespace CSharp.Rendering_and_Printing
{
    class ReadActiveXControlProperties
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

            // Load the documents which store the shapes we want to render.           
            Document doc = new Document(dataDir + "ActiveXControl.docx");

            // Retrieve the target shape from the document. In our sample document this is the first shape.            
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            OleControl oleControl = shape.OleFormat.OleControl;

            string properties = "";
            if (oleControl.IsForms2OleControl)
            {
                Forms2OleControl checkBox = (Forms2OleControl)oleControl;
                properties = "\nCaption: " + checkBox.Caption;
                properties = properties + "\nValue: " + checkBox.Value;
                properties = properties + "\nEnabled: " + checkBox.Enabled;
                properties = properties + "\nType: " + checkBox.Type;
                if (checkBox.ChildNodes != null)
                {
                    properties = properties + "\nChildNodes: " + checkBox.ChildNodes;
                }
            }
            Console.WriteLine("\nActiveX Control properties are " + properties);
        }

        

       
    }
}
