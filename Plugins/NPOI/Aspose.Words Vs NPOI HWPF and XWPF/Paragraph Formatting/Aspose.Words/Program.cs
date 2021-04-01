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
