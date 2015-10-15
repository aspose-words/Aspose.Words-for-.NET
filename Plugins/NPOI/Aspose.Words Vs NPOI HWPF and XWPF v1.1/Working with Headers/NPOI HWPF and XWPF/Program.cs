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
            XWPFDocument wordDocument = new XWPFDocument(new FileStream("data/Working with Headers.doc", FileMode.Open));
            IList<XWPFHeader> headers = wordDocument.HeaderList;
            foreach (XWPFHeader header in headers)
            {
                Console.WriteLine(header.Text);
            }
        }
    }
}
