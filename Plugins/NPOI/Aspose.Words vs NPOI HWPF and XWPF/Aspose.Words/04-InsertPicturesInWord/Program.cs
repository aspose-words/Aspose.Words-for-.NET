using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;

namespace InsertPicturesInWord
{
    class Program
    {
        static void Main(string[] args)
        {
            //----------------------------------------------------
            //  NPOI
            //----------------------------------------------------  
            //const int emusPerInch = 914400;
            //const int emusPerCm = 360000;
            //XWPFDocument doc = new XWPFDocument();
            //XWPFParagraph p2 = doc.CreateParagraph();
            //XWPFRun r2 = p2.CreateRun();
            //r2.SetText("test");

            //var widthEmus = (int)(400.0 * 9525);
            //var heightEmus = (int)(300.0 * 9525);

            //using (FileStream picData = new FileStream("../../image/HumpbackWhale.jpg", FileMode.Open, FileAccess.Read))
            //{
            //    r2.AddPicture(picData, (int)PictureType.PNG, "image1", widthEmus, heightEmus);
            //}
            //using (FileStream sw = File.Create("test.docx"))
            //{
            //    doc.Write(sw);
            //}


            //----------------------------------------------------
            //  Aspose.Words
            //----------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("test");

            var widthEmus = (int)(400.0 * 9525);
            var heightEmus = (int)(300.0 * 9525);

            builder.InsertImage("../../image/HumpbackWhale.jpg", widthEmus, heightEmus);
            doc.Save("test.docx");

        }

    }
}
