using Aspose.Words.Drawing.Charts;
using Aspose.Words;
using NUnit.Framework;

namespace PluginsExamples
{
    public class ProcessorSvgPlugin : PluginsExamplesBase
    {
        [Test]
        public void CreateChartSvg()
        {
            //ExStart:CreateChartSvg
            //GistId:f907ed3ba3d7c0141076152023d37043
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            var shape = builder.InsertChart(ChartType.Pie, 432, 252);
            var chart = shape.Chart;
            chart.Title.Text = "Produced by Aspose.Words Processor plugin.";

            chart.Series.Clear();
            chart.Series.Add("Series 1",
                new string[] { "Category 1", "Category 2", "Category 3" },
                new double[] { 2.7, 3.2, 0.8 });

            doc.Save(ArtifactsDir + "ProcessorSvgPlugin.CreateChartSvg.svg");
            //ExEnd:CreateChartSvg
        }
    }
}
