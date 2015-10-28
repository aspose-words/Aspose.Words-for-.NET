using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI_HWPF_and_XWPF
{
    class Program
    {
        static void Main(string[] args)
        {
            XWPFDocument  wordDocument = new XWPFDocument( new FileStream("data/Convert Word Doc to Other Formats.doc", FileMode.Open));

            using (FileStream sw = File.Create("data/Convert Word Doc to Other Formatsblank.docx"))
            {
                wordDocument.Write(sw);
            }
            
        }
    }
}
