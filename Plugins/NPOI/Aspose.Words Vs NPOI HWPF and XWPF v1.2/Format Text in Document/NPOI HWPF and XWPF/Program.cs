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

            XWPFParagraph p1 = doc.CreateParagraph();
            p1.Alignment = ParagraphAlignment.CENTER;
            p1.BorderBottom = Borders.Double;
            p1.BorderTop = Borders.Double;

            p1.BorderRight = Borders.Double;
            p1.BorderLeft = Borders.Double;
            p1.BorderBetween = Borders.Double;

            p1.VerticalAlignment = TextAlignment.TOP;

            XWPFRun r1 = p1.CreateRun();
            r1.SetText("Hello World");
            r1.IsBold = true;
            r1.FontFamily = "Arial";
            r1.SetUnderline(UnderlinePatterns.DotDotDash);
            r1.SetTextPosition(100);

            using (FileStream sw = File.Create("Formating.docx"))
            {
                doc.Write(sw);
            }
        }
    }
}
