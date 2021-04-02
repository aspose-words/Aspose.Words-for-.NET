using NPOI.XWPF.UserModel;
using System.IO;

namespace NPOI_HWPF_and_XWPF
{
    class Program
    {
        static void Main(string[] args)
        {             

            // Create a new document from scratch
            XWPFDocument doc = new XWPFDocument();
            XWPFTable table = doc.CreateTable(3, 3);

            table.GetRow(1).GetCell(1).SetText("EXAMPLE OF TABLE");

            XWPFTableCell c1 = table.GetRow(0).GetCell(0);
            XWPFParagraph p1 = c1.AddParagraph();   //don't use doc.CreateParagraph
            XWPFRun r1 = p1.CreateRun();
            r1.SetText("This is test table contents");
            r1.SetBold(true);
 
            r1.FontFamily = "Courier";
            r1.SetUnderline(UnderlinePatterns.DotDotDash);
            r1.SetTextPosition(100);
            c1.SetColor("FF0000");


            table.GetRow(2).GetCell(2).SetText("only text");

            FileStream out1 = new FileStream("Format Table in Document.docx", FileMode.Create);
            doc.Write(out1);
            out1.Close();
        }
    }
}
