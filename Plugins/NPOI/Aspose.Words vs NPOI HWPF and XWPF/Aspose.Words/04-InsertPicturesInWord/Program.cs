using System;
using Aspose.Words;
using System.IO;

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

            // Check for license and apply if exists
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                // Apply Aspose.Words API License
                Aspose.Words.License license = new Aspose.Words.License();
                // Place license file in Bin/Debug/ Folder
                license.SetLicense("Aspose.Words.lic");
            }


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
