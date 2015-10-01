using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI_HWPR_and_XWPF
{
    class Program
    {
        static void Main(string[] args)
        {
            XWPFDocument doc = new XWPFDocument(new FileStream("data/Extract Images from Word Document.doc",FileMode.Open));
            IList<XWPFPictureData> pics = doc.AllPictures;

            foreach (XWPFPictureData pic in pics)
            {
                FileStream outputStream = new FileStream("data/NPOI_" + pic.FileName,FileMode.OpenOrCreate);
               byte[] picData= pic.Data;
                outputStream.Write(picData, 0, picData.Length);
                outputStream.Close();
            }
        }
    }
}
