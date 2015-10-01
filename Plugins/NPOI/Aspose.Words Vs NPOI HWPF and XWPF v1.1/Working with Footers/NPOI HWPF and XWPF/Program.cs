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
            XWPFDocument wordDocument = new XWPFDocument(new FileStream("data/Working with Footers.doc", FileMode.Open));
            IList<XWPFFooter> footers = wordDocument.FooterList;
            foreach (XWPFFooter footer in footers)
            {
                Console.WriteLine(footer.Text);
            }
        }
    }
}
