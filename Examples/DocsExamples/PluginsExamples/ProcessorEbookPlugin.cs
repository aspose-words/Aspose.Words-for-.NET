using Aspose.Words.Drawing.Charts;
using Aspose.Words;
using NUnit.Framework;

namespace PluginsExamples
{
    public class ProcessorEbookPlugin : PluginsExamplesBase
    {
        [Test]
        public void CreateChartEbook()
        {
            //ExStart:CreateChartEbook
            //GistId:de897b188462e0289d314d231c3295e8
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            var shape = builder.InsertChart(ChartType.Pie, 432, 252);
            var chart = shape.Chart;
            chart.Title.Text = "Produced by Aspose.Words Processor plugin.";

            chart.Series.Clear();
            chart.Series.Add("Series 1",
                new string[] { "Category 1", "Category 2", "Category 3" },
                new double[] { 2.7, 3.2, 0.8 });

            doc.Save(ArtifactsDir + "ProcessorMSWordPlugin.CreateChartEbook.epub");
            //ExEnd:CreateChartEbook
        }

        [Test]
        public void CreateBookmarkEbook()
        {
            //ExStart:CreateBookmarkEbook
            //GistId:de897b188462e0289d314d231c3295e8
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.StartBookmark("Bookmark");
            builder.Write("Produced by Aspose.Words Processor plugin.");
            builder.EndBookmark("Bookmark");

            doc.Save(ArtifactsDir + "ProcessorMSWordPlugin.CreateBookmarkEbook.epub");
            //ExEnd:CreateBookmarkEbook
        }

        [Test]
        public void EditDocumentEbook()
        {
            //ExStart:EditDocumentEbook
            //GistId:de897b188462e0289d314d231c3295e8
            var doc = new Document(MyDir + "Epub document.epub");
            var builder = new DocumentBuilder(doc);

            builder.MoveToDocumentEnd();
            builder.Writeln("Produced by Aspose.Words Processor plugin.");

            doc.Save(ArtifactsDir + "ProcessorMSWordPlugin.EditDocumentEbook.epub");
            //ExEnd:EditDocumentEbook
        }
    }
}
