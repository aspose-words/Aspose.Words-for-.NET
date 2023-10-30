using Aspose.Words.Drawing.Charts;
using Aspose.Words;
using NUnit.Framework;

namespace PluginsExamples
{
    public class ProcessorXlsxFilesPlugin : PluginsExamplesBase
    {
        [Test]
        public void CreateChartXlsxFiles()
        {
            //ExStart:CreateChartXlsxFiles
            //GistId:e57f464b45000561f7792eef06161c11
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            var shape = builder.InsertChart(ChartType.Pie, 432, 252);
            var chart = shape.Chart;
            chart.Title.Text = "Produced by Aspose.Words Processor plugin.";

            chart.Series.Clear();
            chart.Series.Add("Series 1",
                new string[] { "Category 1", "Category 2", "Category 3" },
                new double[] { 2.7, 3.2, 0.8 });

            doc.Save(ArtifactsDir + "ProcessorXlsxFilesPlugin.CreateChartXlsxFiles.xlsx");
            //ExEnd:CreateChartXlsxFiles
        }
    }
}
