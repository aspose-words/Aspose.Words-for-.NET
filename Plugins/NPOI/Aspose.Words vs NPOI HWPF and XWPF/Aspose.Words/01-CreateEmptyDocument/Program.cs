using System;
using Aspose.Words;
using System.IO;

namespace CreateEmptyDocument
{
    class Program
    {
        static void Main(string[] args)
        {
            //----------------------------------------------------
            //  NPOI
            //----------------------------------------------------            
            //XWPFDocument doc = new XWPFDocument();
            //doc.CreateParagraph();
            //using (FileStream sw = File.Create("blank.docx"))
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
            doc.Save("blank.docx");
        }
    }
}
