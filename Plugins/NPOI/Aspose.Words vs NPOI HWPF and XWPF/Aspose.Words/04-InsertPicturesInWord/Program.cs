using System;
using Aspose.Words;
using System.IO;
using NPOI.XWPF.UserModel;
using Document = Aspose.Words.Document;

namespace InsertPicturesInWord
{
    class Program
    {
        static void Main(string[] args)
        {
            InsertPicturesInWordNPOI();
            InsertPicturesInWordAspose();
        }

        private static void InsertPicturesInWordNPOI()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p2 = doc.CreateParagraph();
            XWPFRun r2 = p2.CreateRun();
            r2.SetText("Hello world!");

            var widthEmus = (int)(400.0 * 9525);
            var heightEmus = (int)(400.0 * 9525);

            using (FileStream picData = new FileStream("../../image/Logo.jpg", FileMode.Open, FileAccess.Read))
            {
                r2.AddPicture(picData, (int)PictureType.PNG, "image1", widthEmus, heightEmus);
            }
            using (FileStream sw = File.Create("InsertPicturesInWordNPOI.docx"))
            {
                doc.Write(sw);
            }
        }

        private static void InsertPicturesInWordAspose()
        {
            // Check for license and apply if exists.
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                // Apply Aspose.Words API License.
                Aspose.Words.License license = new Aspose.Words.License();
                // Place license file in Bin/Debug/ Folder.
                license.SetLicense("Aspose.Words.lic");
            }

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");

            builder.InsertImage("../../image/Logo.jpg", 400, 400);
            doc.Save("InsertPicturesInWordAspose.docx");
        }
    }
}
