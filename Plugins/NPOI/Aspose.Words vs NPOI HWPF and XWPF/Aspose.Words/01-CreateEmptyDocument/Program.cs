using System;
using System.IO;
using NPOI.XWPF.UserModel;
using Document = Aspose.Words.Document;

namespace CreateEmptyDocument
{
    class Program
    {
        static void Main(string[] args)
        {
            CreateEmptyDocumentNPOI();
            CreateEmptyDocumentAspose();
        }

        private static void CreateEmptyDocumentNPOI()
        {
            XWPFDocument doc = new XWPFDocument();
            doc.CreateParagraph();

            using (FileStream sw = File.Create("CreateEmptyDocumentNPOI.docx"))
            {
                doc.Write(sw);
            }
        }

        private static void CreateEmptyDocumentAspose()
        {
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                // Apply Aspose.Words API License
                Aspose.Words.License license = new Aspose.Words.License();
                // Place license file in Bin/Debug/ Folder
                license.SetLicense("Aspose.Words.lic");
            }

            Document doc = new Document();
            doc.Save("CreateEmptyDocumentAspose.docx");
        }
    }
}
