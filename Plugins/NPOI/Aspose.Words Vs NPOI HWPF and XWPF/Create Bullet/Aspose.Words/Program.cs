using System;
using System.IO;

namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
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

            doc.Save("CreateBulletAspose.docx");
        }
    }
}
