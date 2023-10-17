using Aspose.Words.Drawing.Charts;
using Aspose.Words;
using NUnit.Framework;

namespace PluginsExamples
{
    public class ProcessorWebFilesPlugin : PluginsExamplesBase
    {
        [Test]
        public void CreateChartWebFiles()
        {
            //ExStart:CreateChartWebFiles
            //GistId:5b8d06f67f7632970c0d3c3475f998cc
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            var shape = builder.InsertChart(ChartType.Pie, 432, 252);
            var chart = shape.Chart;
            chart.Title.Text = "Produced by Aspose.Words Processor plugin.";

            chart.Series.Clear();
            chart.Series.Add("Series 1",
                new string[] { "Category 1", "Category 2", "Category 3" },
                new double[] { 2.7, 3.2, 0.8 });

            doc.Save(ArtifactsDir + "ProcessorWebFilesPlugin.CreateChartWebFiles.html");
            //ExEnd:CreateChartWebFiles
        }

        [Test]
        public void CreateBookmarkWebFiles()
        {
            //ExStart:CreateBookmarkWebFiles
            //GistId:5b8d06f67f7632970c0d3c3475f998cc
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.StartBookmark("Bookmark");
            builder.Write("Produced by Aspose.Words Processor plugin.");
            builder.EndBookmark("Bookmark");

            doc.Save(ArtifactsDir + "ProcessorWebFilesPlugin.CreateBookmarkWebFiles.html");
            //ExEnd:CreateBookmarkWebFiles
        }

        [Test]
        public void EditDocumentWebFiles()
        {
            //ExStart:EditDocumentWebFiles
            //GistId:5b8d06f67f7632970c0d3c3475f998cc
            var doc = new Document(MyDir + "Document.html");
            var builder = new DocumentBuilder(doc);

            builder.MoveToDocumentEnd();
            builder.Writeln("Produced by Aspose.Words Processor plugin.");

            doc.Save(ArtifactsDir + "ProcessorWebFilesPlugin.EditDocumentWebFiles.html");
            //ExEnd:EditDocumentWebFiles
        }
    }
}
