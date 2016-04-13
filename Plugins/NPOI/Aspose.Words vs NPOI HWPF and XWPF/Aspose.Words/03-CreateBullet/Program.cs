using System;
using Aspose.Words;
using System.IO;

namespace CreateBullet
{
    class Program
    {
        static void Main(string[] args)
        {
            //----------------------------------------------------
            //  NPOI
            //----------------------------------------------------  
            //XWPFDocument doc = new XWPFDocument();
            ////simple bullet
            //XWPFNumbering numbering = doc.CreateNumbering();
            
            //string abstractNumId = numbering.AddAbstractNum();
            //string numId = numbering.AddNum(abstractNumId);

            //XWPFParagraph p0 = doc.CreateParagraph();
            //XWPFRun r0 = p0.CreateRun();
            //r0.SetText("simple bullet");
            //r0.SetBold(true);
            //r0.FontFamily = "Courier";
            //r0.FontSize = 12;

            //XWPFParagraph p1 = doc.CreateParagraph();
            //XWPFRun r1 = p1.CreateRun();
            //r1.SetText("first, create paragraph and run, set text");
            //p1.SetNumID(numId);

            //XWPFParagraph p2 = doc.CreateParagraph();
            //XWPFRun r2 = p2.CreateRun();
            //r2.SetText("second, call XWPFDocument.CreateNumbering() to create numbering");
            //p2.SetNumID(numId);

            //XWPFParagraph p3 = doc.CreateParagraph();
            //XWPFRun r3 = p3.CreateRun();
            //r3.SetText("third, add AbstractNum[numbering.AddAbstractNum()] and Num(numbering.AddNum(abstractNumId))");
            //p3.SetNumID(numId);

            //XWPFParagraph p4 = doc.CreateParagraph();
            //XWPFRun r4 = p4.CreateRun();
            //r4.SetText("next, call XWPFParagraph.SetNumID(numId) to set paragraph property, CT_P.pPr.numPr");
            //p4.SetNumID(numId);

            ////multi level
            //abstractNumId = numbering.AddAbstractNum();
            //numId = numbering.AddNum(abstractNumId);
            //doc.CreateParagraph();
            //doc.CreateParagraph();

            //p1 = doc.CreateParagraph();
            //r1 = p1.CreateRun();
            //r1.SetText("multi level bullet");
            //r1.SetBold(true);
            //r1.FontFamily = "Courier";
            //r1.FontSize =12 ;

            //p1 = doc.CreateParagraph();
            //r1 = p1.CreateRun();
            //r1.SetText("first");
            //p1.SetNumID(numId, "0");
            //p1 = doc.CreateParagraph();
            //r1 = p1.CreateRun();
            //r1.SetText("first-first");
            //p1.SetNumID(numId, "1");
            //p1 = doc.CreateParagraph();
            //r1 = p1.CreateRun();
            //r1.SetText("first-second");
            //p1.SetNumID(numId, "1");
            //p1 = doc.CreateParagraph();
            //r1 = p1.CreateRun();
            //r1.SetText("first-third");
            //p1.SetNumID(numId, "1");
            //p1 = doc.CreateParagraph();
            //r1 = p1.CreateRun();
            //r1.SetText("second");
            //p1.SetNumID(numId, "0");
            //p1 = doc.CreateParagraph();
            //r1 = p1.CreateRun();
            //r1.SetText("second-first");
            //p1.SetNumID(numId, "1");
            //p1 = doc.CreateParagraph();
            //r1 = p1.CreateRun();
            //r1.SetText("second-second");
            //p1.SetNumID(numId, "1");
            //p1 = doc.CreateParagraph();
            //r1 = p1.CreateRun();
            //r1.SetText("second-third");
            //p1.SetNumID(numId, "1");
            //p1 = doc.CreateParagraph();
            //r1 = p1.CreateRun();
            //r1.SetText("second-third-first");
            //p1.SetNumID(numId, "2");
            //p1 = doc.CreateParagraph();
            //r1 = p1.CreateRun();
            //r1.SetText("second-third-second");
            //p1.SetNumID(numId, "2");
            //p1 = doc.CreateParagraph();
            //r1 = p1.CreateRun();
            //r1.SetText("third");
            //p1.SetNumID(numId, "0");

            //FileStream sw = new FileStream("bullet-sample.docx", FileMode.Create);
            //doc.Write(sw);
            //sw.Close();

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
            
            builder.Font.Bold = true;
            builder.Font.Name = "Courier";
            builder.Font.Size = 12;
            builder.Writeln("simple bullet");
            builder.Font.ClearFormatting();
            builder.ListFormat.List = doc.Lists.Add(Aspose.Words.Lists.ListTemplate.BulletSquare);
            builder.Writeln("first, create paragraph and run, set text");
            builder.Writeln("second, call XWPFDocument.CreateNumbering() to create numbering");
            builder.Writeln("third, add AbstractNum[numbering.AddAbstractNum()] and Num(numbering.AddNum(abstractNumId))");
            builder.Writeln("next, call XWPFParagraph.SetNumID(numId) to set paragraph property, CT_P.pPr.numPr");
            builder.ListFormat.RemoveNumbers();
            builder.InsertBreak(BreakType.ParagraphBreak);

            //multi level
            builder.Font.Bold = true;
            builder.Font.Name = "Courier";
            builder.Font.Size = 12;
            builder.Writeln("multi level bullet");
            builder.Font.ClearFormatting();
            builder.ListFormat.List = doc.Lists.Add(Aspose.Words.Lists.ListTemplate.BulletSquare);
            builder.Writeln("first");
            builder.ListFormat.ListLevelNumber = 1;
            builder.Writeln("first-first");
            builder.Writeln("first-second");
            builder.Writeln("first-third");
            builder.ListFormat.List = doc.Lists.Add(Aspose.Words.Lists.ListTemplate.BulletSquare);            
            builder.ListFormat.ListLevelNumber = 0;
            builder.Writeln("second");
            builder.ListFormat.ListLevelNumber = 1;
            builder.Writeln("second-first");
            builder.Writeln("second-second");
            builder.Writeln("second-third");
            builder.ListFormat.ListLevelNumber = 2;
            builder.Writeln("second-third-first");
            builder.Writeln("second-third-second");
            builder.ListFormat.List = doc.Lists.Add(Aspose.Words.Lists.ListTemplate.BulletSquare);
            builder.ListFormat.ListLevelNumber = 0;            
            builder.Writeln("third");
            builder.ListFormat.RemoveNumbers();                      

            doc.Save("bullet-sample.docx");

        }

    }
}
