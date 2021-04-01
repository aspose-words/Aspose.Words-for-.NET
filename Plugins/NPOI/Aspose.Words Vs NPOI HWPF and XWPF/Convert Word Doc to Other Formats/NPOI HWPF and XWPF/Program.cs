using System.IO;
using NPOI.XWPF.UserModel;

namespace NPOI_HWPF_and_XWPF
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = Path.GetDirectoryName(Path.GetDirectoryName(Directory.GetCurrentDirectory())) + @"\data\" + "Convert Word Doc to Other Formats.doc";
            XWPFDocument wordDocument = new XWPFDocument();
            FileStream out1 = new FileStream(filePath, FileMode.Open); 
            using (FileStream sw = File.Create("Convert Word Doc to Other Formatsblank.docx"))
            {               
                wordDocument.Write(sw);
            }
            
        }
    }
}
