using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
            Document doc = new Document();
            doc.Save("blank.docx");
        }
    }
}
