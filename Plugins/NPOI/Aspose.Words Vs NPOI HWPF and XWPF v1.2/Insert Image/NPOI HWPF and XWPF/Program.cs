using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NPOI_HWPF_and_XWPF
{
    class Program
    {
        static void Main(string[] args)
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p = doc.CreateParagraph();

            XWPFRun r = p.CreateRun();
            
            
            
            FileStream fileStream = new FileStream("Image.png", FileMode.Open, FileAccess.Read);
            r.AddPicture(fileStream,2, "Image.png",100,100);

            using (FileStream sw = File.Create("Insert Image.docx"))
            {
                doc.Write(sw);
            }

        }
    }
}
