using Aspose.Words.Drawing.Charts;
using Aspose.Words;
using NUnit.Framework;
using Aspose.Words.Tables;
using System.Drawing;

namespace PluginsExamples
{
    public class ProcessorXlsxFilesPlugin : PluginsExamplesBase
    {
        [Test]
        public void CreateTableXlsxFiles()
        {
            //ExStart:CreateTableXlsxFiles
            //GistId:e57f464b45000561f7792eef06161c11
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            
            builder.StartTable();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;            
            builder.CellFormat.Shading.BackgroundPatternColor = Color.AliceBlue;

            for (int i = 0; i < 3; i++)
            {
                builder.InsertCell();
                builder.Write($"Row {i + 1}, Column 1");
                builder.InsertCell();
                builder.Write($"Row {i + 1}, Column 2");

                Row row = builder.EndRow();

                builder.CellFormat.Shading.ClearFormatting();

                BorderCollection borders = row.RowFormat.Borders;
                // Adjust the appearance of borders that will appear between rows.
                borders.Horizontal.Color = Color.Red;
                borders.Horizontal.LineStyle = LineStyle.Dot;
                borders.Horizontal.LineWidth = 2.0d;
                // Adjust the appearance of borders that will appear between cells.
                borders.Vertical.Color = Color.Blue;
                borders.Vertical.LineStyle = LineStyle.Dot;
                borders.Vertical.LineWidth = 2.0d;
            }

            doc.Save(ArtifactsDir + "ProcessorXlsxFilesPlugin.CreateTableXlsxFiles.xlsx");
            //ExEnd:CreateTableXlsxFiles
        }
    }
}
