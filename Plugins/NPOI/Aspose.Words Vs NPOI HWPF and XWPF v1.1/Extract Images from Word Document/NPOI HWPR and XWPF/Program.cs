using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace NPOI_HWPR_and_XWPF
{
    class Program
    {
        static void Main(string[] args)
        {

            string filePath = Path.GetDirectoryName(Path.GetDirectoryName(Directory.GetCurrentDirectory())) + @"\data\" + "Extract Images from Word Document.doc";

            // NPOI library doest not have ablitity to read word document. 
            XWPFDocument doc = new XWPFDocument(new FileStream(filePath, FileMode.Open));
            IList<XWPFPictureData> pics = doc.AllPictures;


            foreach (XWPFPictureData pic in pics)
            {
                FileStream outputStream = new FileStream("data/NPOI_" + pic.FileName, FileMode.OpenOrCreate);
                byte[] picData = pic.Data;
                outputStream.Write(picData, 0, picData.Length);
                outputStream.Close();
            }
        }
    }
}
