using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Ole;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents
{
    class WorkingWithOleObjectsAndActiveX : DocsExamplesBase
    {
        [Test]
        public void InsertOleObject()
        {
            //ExStart:DocumentBuilderInsertOleObject
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertOleObject("http://www.aspose.com", "htmlfile", true, true, null);

            doc.Save(ArtifactsDir + "WorkingWithOleObjectsAndActiveX.InsertOleObject.docx");
            //ExEnd:DocumentBuilderInsertOleObject
        }

        [Test]
        public void InsertOleObjectWithOlePackage()
        {
            //ExStart:InsertOleObjectwithOlePackage
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            byte[] bs = File.ReadAllBytes(MyDir + "Zip file.zip");
            using (Stream stream = new MemoryStream(bs))
            {
                Shape shape = builder.InsertOleObject(stream, "Package", true, null);
                OlePackage olePackage = shape.OleFormat.OlePackage;
                olePackage.FileName = "filename.zip";
                olePackage.DisplayName = "displayname.zip";

                doc.Save(ArtifactsDir + "WorkingWithOleObjectsAndActiveX.InsertOleObjectWithOlePackage.docx");
            }
            //ExEnd:InsertOleObjectwithOlePackage

            //ExStart:GetAccessToOLEObjectRawData
            Shape oleShape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            byte[] oleRawData = oleShape.OleFormat.GetRawData();
            //ExEnd:GetAccessToOLEObjectRawData
        }

        [Test]
        public void InsertOleObjectAsIcon()
        {
            //ExStart:InsertOLEObjectAsIcon
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertOleObjectAsIcon(MyDir + "Presentation.pptx", false, ImagesDir + "Logo icon.ico",
                "My embedded file");

            doc.Save(ArtifactsDir + "WorkingWithOleObjectsAndActiveX.InsertOleObjectAsIcon.docx");
            //ExEnd:InsertOLEObjectAsIcon
        }

        [Test]
        public void InsertOleObjectAsIconUsingStream()
        {
            //ExStart:InsertOLEObjectAsIconUsingStream
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            using (MemoryStream stream = new MemoryStream(File.ReadAllBytes(MyDir + "Presentation.pptx")))
                builder.InsertOleObjectAsIcon(stream, "Package", ImagesDir + "Logo icon.ico", "My embedded file");

            doc.Save(ArtifactsDir + "WorkingWithOleObjectsAndActiveX.InsertOleObjectAsIconUsingStream.docx");
            //ExEnd:InsertOLEObjectAsIconUsingStream
        }

        [Test]
        public void ReadActiveXControlProperties()
        {
            Document doc = new Document(MyDir + "ActiveX controls.docx");

            string properties = "";
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                if (shape.OleFormat is null) break;

                OleControl oleControl = shape.OleFormat.OleControl;
                if (oleControl.IsForms2OleControl)
                {
                    Forms2OleControl checkBox = (Forms2OleControl) oleControl;
                    properties = properties + "\nCaption: " + checkBox.Caption;
                    properties = properties + "\nValue: " + checkBox.Value;
                    properties = properties + "\nEnabled: " + checkBox.Enabled;
                    properties = properties + "\nType: " + checkBox.Type;
                    if (checkBox.ChildNodes != null)
                    {
                        properties = properties + "\nChildNodes: " + checkBox.ChildNodes;
                    }

                    properties += "\n";
                }
            }

            properties = properties + "\nTotal ActiveX Controls found: " + doc.GetChildNodes(NodeType.Shape, true).Count;
            Console.WriteLine("\n" + properties);
        }
    }
}