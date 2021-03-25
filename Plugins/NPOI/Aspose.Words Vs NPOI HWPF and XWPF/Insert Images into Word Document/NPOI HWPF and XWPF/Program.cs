using System.IO;
using NPOI.XWPF.UserModel;

namespace NPOI_HWPF_and_XWPF
{
    class Program
    {
        static void Main(string[] args)
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
    }
}
