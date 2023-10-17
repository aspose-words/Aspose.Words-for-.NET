using Aspose.Words.Drawing.Charts;
using Aspose.Words;
using NUnit.Framework;
using Aspose.Words.Saving;

namespace PluginsExamples
{
    public class ProcessorPdfPlugin : PluginsExamplesBase
    {
        [Test]
        public void CreateChartPdf()
        {
            //ExStart:CreateChartPdf
            //GistId:f27861863b61d1fb3b986e9450f53389
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            var shape = builder.InsertChart(ChartType.Pie, 432, 252);
            var chart = shape.Chart;
            chart.Title.Text = "Produced by Aspose.Words Processor plugin.";

            chart.Series.Clear();
            chart.Series.Add("Series 1",
                new string[] { "Category 1", "Category 2", "Category 3" },
                new double[] { 2.7, 3.2, 0.8 });

            doc.Save(ArtifactsDir + "ProcessorPdfPlugin.CreateChartPdf.pdf");
            //ExEnd:CreateChartPdf
        }

        [Test]
        public void CreateBookmarkPdf()
        {
            //ExStart:CreateBookmarkPdf
            //GistId:f27861863b61d1fb3b986e9450f53389
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.StartBookmark("Bookmark");
            builder.Write("Produced by Aspose.Words Processor plugin.");
            builder.EndBookmark("Bookmark");

            var saveOptions = new PdfSaveOptions();
            saveOptions.OutlineOptions.BookmarksOutlineLevels.Add("Bookmark", 1);

            doc.Save(ArtifactsDir + "ProcessorPdfPlugin.CreateBookmarkPdf.pdf", saveOptions);
            //ExEnd:CreateBookmarkPdf
        }

        [Test]
        public void EditDocumentPdf()
        {
            //ExStart:EditDocumentPdf
            //GistId:f27861863b61d1fb3b986e9450f53389
            var doc = new Document(MyDir + "Pdf Document.pdf");
            var builder = new DocumentBuilder(doc);

            builder.MoveToDocumentEnd();
            builder.Writeln("Produced by Aspose.Words Processor plugin.");

            doc.Save(ArtifactsDir + "ProcessorPdfPlugin.EditDocumentPdf.pdf");
            //ExEnd:EditDocumentPdf
        }
    }
}
