using System;
using Aspose.Words;
using System.IO;

namespace SimpleTable
{
    class Program
    {
        static void Main(string[] args)
        {
            //----------------------------------------------------
            //  NPOI
            //----------------------------------------------------            
            //XWPFDocument doc = new XWPFDocument();
            //XWPFParagraph para = doc.CreateParagraph();
            //XWPFRun r0 = para.CreateRun();
            //r0.SetText("Title1");
            //para.BorderTop = Borders.THICK;
            //para.FillBackgroundColor = "EEEEEE";
            //para.FillPattern = NPOI.OpenXmlFormats.Wordprocessing.ST_Shd.diagStripe;

            //XWPFTable table = doc.CreateTable(3, 3);

            //table.GetRow(1).GetCell(1).SetText("EXAMPLE OF TABLE");

            //XWPFTableCell c1 = table.GetRow(0).GetCell(0);
            //XWPFParagraph p1 = c1.AddParagraph();   //don't use doc.CreateParagraph
            //XWPFRun r1 = p1.CreateRun();
            //r1.SetText("The quick brown fox");
            //r1.SetBold(true);

            //r1.FontFamily = "Courier";
            //r1.SetUnderline(UnderlinePatterns.DotDotDash);
            //r1.SetTextPosition(100);
            //c1.SetColor("FF0000");


            //table.GetRow(2).GetCell(2).SetText("only text");

            //FileStream out1 = new FileStream("simpleTable.docx", FileMode.Create);
            //doc.Write(out1);
            //out1.Close();


            //----------------------------------------------------
            //  Aspose.Words
            //----------------------------------------------------

            // Check for license and apply if exists
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                // Apply Aspose.Words API License
                Aspose.Words.License license = new Aspose.Words.License();
                // Place license file in Bin/Debug/ Folder
                license.SetLicense("Aspose.Words.lic");
            }


            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ParagraphFormat.Borders.Top.LineStyle = LineStyle.Thick;
            builder.ParagraphFormat.Shading.BackgroundPatternColor = System.Drawing.ColorTranslator.FromHtml("#EEEEEE");
            builder.ParagraphFormat.Shading.Texture = TextureIndex.TextureDarkDiagonalUp;
            builder.Writeln("Title1");
            
            builder.ParagraphFormat.ClearFormatting();            
            builder.InsertBreak(BreakType.ParagraphBreak);

            // We call this method to start building the table.
            builder.StartTable();
            builder.InsertCell();

            builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.ColorTranslator.FromHtml("#FF0000");
            builder.Font.Position = 100;
            builder.Font.Name = "Courier";
            builder.Font.Bold = true;
            builder.Font.Underline = Underline.DotDotDash;
            builder.Write("The quick brown fox");
            builder.InsertCell();

            builder.Font.ClearFormatting();
            builder.CellFormat.ClearFormatting();

            builder.InsertCell();
            builder.EndRow();

            builder.InsertCell();
            builder.InsertCell();
            builder.Write("EXAMPLE OF TABLE");
            builder.InsertCell(); 
            builder.EndRow();

            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.Write("only text");
            builder.EndRow();

            // Signal that we have finished building the table.
            builder.EndTable();

            doc.Save("simpleTable.docx");

        }
    }
}
