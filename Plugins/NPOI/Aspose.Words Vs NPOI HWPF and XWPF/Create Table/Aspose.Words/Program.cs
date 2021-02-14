using System;
using System.IO;

namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            // Check for license, and apply if it exists.
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                // Apply Aspose.Words API License.
                Aspose.Words.License license = new Aspose.Words.License();
                // Place license file in Bin/Debug/ Folder.
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

            // Call this method to start building the table.
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

            doc.Save("SimpleTableAspose.docx");
        }
    }
}
