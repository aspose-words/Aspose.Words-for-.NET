using Aspose.Words;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CSharp.Programming_Documents.Video
{
    class InsertOnlineVideoWithEmbedHtml
    {
        public static void Run()
        {
            //ExStart:InsertOnlineVideoWithEmbedHtml
            //The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithOnlineVideo();

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Shape width/height.
            double width = 360;
            double height = 270;

            // Poster frame image.
            byte[] imageBytes = File.ReadAllBytes("TestImage.jpg");

            // Visible url
            string vimeoVideoUrl = @"https://vimeo.com/52477838";

            // Embed Html code.
            string vimeoEmbedCode = "";

            builder.InsertOnlineVideo(vimeoVideoUrl, vimeoEmbedCode, imageBytes, width, height);

            dataDir = dataDir + "Insert.OnlineVideo_out_.docx";
            doc.Save(dataDir);
            //ExEnd:InsertOnlineVideoWithEmbedHtml
            Console.WriteLine("\nOnline video inserted successfully.\nFile saved at " + dataDir);
        }
    }
}
