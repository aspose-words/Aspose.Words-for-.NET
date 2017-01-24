using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CSharp.Programming_Documents.Video
{
    class InsertOnlineVideo
    {
        public static void Run()
        {
            //ExStart:InsertOnlineVideo
            //The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithOnlineVideo();

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Pass direct url from youtu.be.
            string url = "https://youtu.be/t_1LYZ102RA";

            double width = 360;
            double height = 270;

            Shape shape = builder.InsertOnlineVideo(url, width, height);

            dataDir = dataDir + "Insert.OnlineVideo_out_.docx";
            doc.Save(dataDir);
            //ExEnd:InsertOnlineVideo
            Console.WriteLine("\nOnline video inserted successfully.\nFile saved at " + dataDir);
        }
    }
}
