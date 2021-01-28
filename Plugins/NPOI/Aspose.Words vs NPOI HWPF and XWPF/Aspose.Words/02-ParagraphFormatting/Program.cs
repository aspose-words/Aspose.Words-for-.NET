using System;
using Aspose.Words;
using System.IO;
using NPOI.XWPF.UserModel;
using BreakType = Aspose.Words.BreakType;
using Document = Aspose.Words.Document;
using LineSpacingRule = Aspose.Words.LineSpacingRule;
using ParagraphAlignment = Aspose.Words.ParagraphAlignment;

namespace SimpleDocument
{
    class Program
    {
        static void Main(string[] args)
        {
            ParagraphFormattingNPOI();
            ParagraphFormattingAspose();
        }

        private static void ParagraphFormattingNPOI()
        {
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph p1 = doc.CreateParagraph();
            p1.Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.CENTER;
            p1.BorderBottom = Borders.Double;
            p1.BorderTop = Borders.Double;

            p1.BorderRight = Borders.Double;
            p1.BorderLeft = Borders.Double;
            p1.BorderBetween = Borders.Single;

            p1.VerticalAlignment = TextAlignment.TOP;

            XWPFRun r1 = p1.CreateRun();
            r1.SetText("The quick brown fox");
            r1.IsBold = true;
            r1.FontFamily = "Courier";
            r1.SetUnderline(UnderlinePatterns.DotDotDash);
            r1.TextPosition = 100;

            XWPFParagraph p2 = doc.CreateParagraph();
            p2.Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.RIGHT;

            //BORDERS
            p2.BorderBottom = Borders.Double;
            p2.BorderTop = Borders.Double;
            p2.BorderRight = Borders.Double;
            p2.BorderLeft = Borders.Double;
            p2.BorderBetween = Borders.Single;

            XWPFRun r2 = p2.CreateRun();
            r2.SetText("jumped over the lazy dog");
            r2.IsStrikeThrough = true;
            r2.FontSize = 20;

            XWPFRun r3 = p2.CreateRun();
            r3.SetText("and went away");
            r3.IsStrikeThrough = true;
            r3.FontSize = 20;
            r3.Subscript = VerticalAlign.SUPERSCRIPT;
            r3.SetColor("FF0000");

            XWPFParagraph p3 = doc.CreateParagraph();
            p3.IsWordWrapped = true;
            p3.IsPageBreak = true;
            p3.Alignment = NPOI.XWPF.UserModel.ParagraphAlignment.BOTH;
            p3.SpacingLineRule = NPOI.XWPF.UserModel.LineSpacingRule.EXACT;
            p3.IndentationFirstLine = 600;

            XWPFRun r4 = p3.CreateRun();
            r4.TextPosition = 20;
            r4.SetText("To be, or not to be: that is the question: "
                    + "Whether 'tis nobler in the mind to suffer "
                    + "The slings and arrows of outrageous fortune, "
                    + "Or to take arms against a sea of troubles, "
                    + "And by opposing end them? To die: to sleep; ");
            r4.AddBreak(NPOI.XWPF.UserModel.BreakType.PAGE);
            r4.SetText("No more; and by a sleep to say we end "
                    + "The heart-ache and the thousand natural shocks "
                    + "That flesh is heir to, 'tis a consummation "
                    + "Devoutly to be wish'd. To die, to sleep; "
                    + "To sleep: perchance to dream: ay, there's the rub; "
                    + ".......");
            r4.IsItalic = true;
            //This would imply that this break shall be treated as a simple line break, and break the line after that word:

            XWPFRun r5 = p3.CreateRun();
            r5.TextPosition = -10;
            r5.SetText("For in that sleep of death what dreams may come");
            r5.AddCarriageReturn();
            r5.SetText("When we have shuffled off this mortal coil,"
                    + "Must give us pause: there's the respect"
                    + "That makes calamity of so long life;");
            r5.AddBreak();
            r5.SetText("For who would bear the whips and scorns of time,"
                    + "The oppressor's wrong, the proud man's contumely,");

            r5.AddBreak(BreakClear.ALL);
            r5.SetText("The pangs of despised love, the law's delay,"
                    + "The insolence of office and the spurns" + ".......");

            using (FileStream outDocStream = new FileStream("ParagraphFormattingNPOI.docx", FileMode.Create))
            {
                doc.Write(outDocStream);
            }
        }

        private static void ParagraphFormattingAspose()
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

            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.ParagraphFormat.Borders.LineStyle = LineStyle.Double;
            builder.ParagraphFormat.Borders.Bottom.LineStyle = LineStyle.Single;

            builder.Font.Bold = true;
            builder.Font.Name = "Courier";
            builder.Font.Underline = Underline.DotDotDash;
            builder.Font.Position = 100;
            builder.Writeln("The quick brown fox");

            builder.ParagraphFormat.ClearFormatting();
            builder.ParagraphFormat.Borders.LineStyle = LineStyle.Double;
            builder.ParagraphFormat.Borders.Top.LineStyle = LineStyle.None;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            builder.Font.ClearFormatting();
            builder.Font.StrikeThrough = true;
            builder.Font.Size = 20;
            builder.Write("jumped over the lazy dog");

            builder.Font.StrikeThrough = true;
            builder.Font.Size = 20;
            builder.Font.Superscript = true;
            builder.Font.Color = System.Drawing.ColorTranslator.FromHtml("#FF0000");
            builder.Writeln("and went away");
            builder.ParagraphFormat.ClearFormatting();

            builder.InsertBreak(BreakType.PageBreak);
            builder.Font.ClearFormatting();

            builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
            builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Exactly;
            builder.ParagraphFormat.FirstLineIndent = 600;

            builder.Font.Position = 20;
            builder.Writeln("To be, or not to be: that is the question: "
                            + "Whether 'tis nobler in the mind to suffer "
                            + "The slings and arrows of outrageous fortune, "
                            + "Or to take arms against a sea of troubles, "
                            + "And by opposing end them? To die: to sleep; ");

            builder.ParagraphFormat.ClearFormatting();
            builder.Font.ClearFormatting();
            builder.Font.Italic = true;
            builder.Writeln("No more; and by a sleep to say we end "
                            + "The heart-ache and the thousand natural shocks "
                            + "That flesh is heir to, 'tis a consummation "
                            + "Devoutly to be wish'd. To die, to sleep; "
                            + "To sleep: perchance to dream: ay, there's the rub; "
                            + ".......");

            builder.ParagraphFormat.ClearFormatting();
            builder.Font.ClearFormatting();

            builder.Font.Position = 10;
            builder.Writeln("For in that sleep of death what dreams may come");
            builder.InsertBreak(BreakType.LineBreak);
            builder.Writeln("When we have shuffled off this mortal coil,"
                            + "Must give us pause: there's the respect"
                            + "That makes calamity of so long life;");

            builder.InsertBreak(BreakType.LineBreak);

            builder.Writeln("For who would bear the whips and scorns of time,"
                            + "The oppressor's wrong, the proud man's contumely,");

            builder.InsertBreak(BreakType.LineBreak);

            builder.Font.ClearFormatting();
            builder.Writeln("The pangs of despised love, the law's delay,"
                            + "The insolence of office and the spurns" + ".......");

            doc.Save("ParagraphFormattingAspose.docx");
        }
    }
}
