using Aspose.Words.Drawing.Charts;
using Aspose.Words;
using NUnit.Framework;

namespace PluginsExamples
{
    public class ProcessorOpenOfficePlugin : PluginsExamplesBase
    {
        [Test]
        public void CreateChartOpenOffice()
        {
            //ExStart:CreateChartOpenOffice
            //GistId:65aba613db784994b7b8d285fdc37433
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            var shape = builder.InsertChart(ChartType.Pie, 432, 252);
            var chart = shape.Chart;
            chart.Title.Text = "Produced by Aspose.Words Processor plugin.";

            chart.Series.Clear();
            chart.Series.Add("Series 1",
                new string[] { "Category 1", "Category 2", "Category 3" },
                new double[] { 2.7, 3.2, 0.8 });

            doc.Save(ArtifactsDir + "ProcessorOpenOfficePlugin.CreateChartOpenOffice.odt");
            //ExEnd:CreateChartOpenOffice
        }

        [Test]
        public void CreateBookmarkOpenOffice()
        {
            //ExStart:CreateBookmarkOpenOffice
            //GistId:65aba613db784994b7b8d285fdc37433
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.StartBookmark("Bookmark");
            builder.Write("Produced by Aspose.Words Processor plugin.");
            builder.EndBookmark("Bookmark");

            doc.Save(ArtifactsDir + "ProcessorOpenOfficePlugin.CreateBookmarkOpenOffice.odt");
            //ExEnd:CreateBookmarkOpenOffice
        }

        [Test]
        public void EditDocumentOpenOffice()
        {
            //ExStart:EditDocumentOpenOffice
            //GistId:65aba613db784994b7b8d285fdc37433
            var doc = new Document(MyDir + "Document.odt");
            var builder = new DocumentBuilder(doc);

            builder.MoveToDocumentEnd();
            builder.Writeln("Produced by Aspose.Words Processor plugin.");

            doc.Save(ArtifactsDir + "ProcessorOpenOfficePlugin.EditDocumentOpenOffice.odt");
            //ExEnd:EditDocumentOpenOffice
        }
    }
}
