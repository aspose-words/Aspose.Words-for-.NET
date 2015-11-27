using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
            
            builder.Font.Bold = true;
            builder.Font.Name = "Arial";            
            builder.Font.Position = 100;
            builder.Writeln("Hello World");

            builder.ParagraphFormat.ClearFormatting();
            builder.ParagraphFormat.Borders.LineStyle = LineStyle.Double;
            builder.ParagraphFormat.Borders.Top.LineStyle = LineStyle.None;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            builder.Font.ClearFormatting();
            
            builder.Font.Size = 25;
            builder.Write("Hello Aspose");

           

            doc.Save("Formating.docx");
        }
    }
}
