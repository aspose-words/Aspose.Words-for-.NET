using System;
using System.IO;
using System.Runtime.InteropServices.ComTypes;
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
            //ExStart:InsertOleObject
            //GistId:4996b573cf231d9f66ab0d1f3f981222
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertOleObject("http://www.aspose.com", "htmlfile", true, true, null);

            doc.Save(ArtifactsDir + "WorkingWithOleObjectsAndActiveX.InsertOleObject.docx");
            //ExEnd:InsertOleObject
        }

        [Test]
        public void InsertOleObjectWithOlePackage()
        {
            //ExStart:InsertOleObjectwithOlePackage
            //GistId:4996b573cf231d9f66ab0d1f3f981222
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

            //ExStart:GetAccessToOleObjectRawData
            //GistId:4996b573cf231d9f66ab0d1f3f981222
            Shape oleShape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            byte[] oleRawData = oleShape.OleFormat.GetRawData();
            //ExEnd:GetAccessToOleObjectRawData
        }

        [Test]
        public void InsertOleObjectAsIcon()
        {
            //ExStart:InsertOleObjectAsIcon
            //GistId:4996b573cf231d9f66ab0d1f3f981222
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertOleObjectAsIcon(MyDir + "Presentation.pptx", false, ImagesDir + "Logo icon.ico",
                "My embedded file");

            doc.Save(ArtifactsDir + "WorkingWithOleObjectsAndActiveX.InsertOleObjectAsIcon.docx");
            //ExEnd:InsertOleObjectAsIcon
        }

        [Test]
        public void InsertOleObjectAsIconUsingStream()
        {
            //ExStart:InsertOleObjectAsIconUsingStream
            //GistId:4996b573cf231d9f66ab0d1f3f981222
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            using (MemoryStream stream = new MemoryStream(File.ReadAllBytes(MyDir + "Presentation.pptx")))
                builder.InsertOleObjectAsIcon(stream, "Package", ImagesDir + "Logo icon.ico", "My embedded file");

            doc.Save(ArtifactsDir + "WorkingWithOleObjectsAndActiveX.InsertOleObjectAsIconUsingStream.docx");
            //ExEnd:InsertOleObjectAsIconUsingStream
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

        [Test]
        public void InsertOnlineVideo()
        {
            //ExStart:InsertOnlineVideo
            //GistId:4996b573cf231d9f66ab0d1f3f981222
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            string url = "https://youtu.be/t_1LYZ102RA";
            double width = 360;
            double height = 270;

            Shape shape = builder.InsertOnlineVideo(url, width, height);

            doc.Save(ArtifactsDir + "WorkingWithOleObjectsAndActiveX.InsertOnlineVideo.docx");
            //ExEnd:InsertOnlineVideo
        }

        [Test]
        public void InsertOnlineVideoWithEmbedHtml()
        {
            //ExStart:InsertOnlineVideoWithEmbedHtml
            //GistId:4996b573cf231d9f66ab0d1f3f981222
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            double width = 360;
            double height = 270;

            string videoUrl = "https://vimeo.com/52477838";
            string videoEmbedCode =
                "<iframe src=\"https://player.vimeo.com/video/52477838\" width=\"640\" height=\"360\" frameborder=\"0\" " +
                "title=\"Aspose\" webkitallowfullscreen mozallowfullscreen allowfullscreen></iframe>";

            byte[] thumbnailImageBytes = File.ReadAllBytes(ImagesDir + "Logo.jpg");

            builder.InsertOnlineVideo(videoUrl, videoEmbedCode, thumbnailImageBytes, width, height);

            doc.Save(ArtifactsDir + "WorkingWithOleObjectsAndActiveX.InsertOnlineVideoWithEmbedHtml.docx");
            //ExEnd:InsertOnlineVideoWithEmbedHtml
        }
    }
}