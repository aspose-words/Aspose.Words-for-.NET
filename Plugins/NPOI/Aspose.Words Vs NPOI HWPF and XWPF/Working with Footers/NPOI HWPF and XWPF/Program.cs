using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace NPOI_HWPF_and_XWPF
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = Path.GetDirectoryName(Path.GetDirectoryName(Directory.GetCurrentDirectory())) + @"\data\" + "Footers.doc";
            // NPOI library doest not have ablitity to read word document. 
            XWPFDocument wordDocument = new XWPFDocument(new FileStream(filePath, FileMode.Open));
            IList<XWPFFooter> footers = wordDocument.FooterList;
            foreach (XWPFFooter footer in footers)
            {
                Console.WriteLine(footer.Text);
            }
        }
    }
}
