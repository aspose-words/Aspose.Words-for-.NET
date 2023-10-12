using Aspose.Words.Drawing.Charts;
using Aspose.Words;
using NUnit.Framework;

namespace PluginsExamples
{
    public class ProcessorWordPlugin : PluginsExamplesBase
    {
        [Test]
        public void CreateChartWord()
        {
            //ExStart:CreateChartWord
            //GistId:c3b9534ddda2dbbd6f267b53dda5b605
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            var shape = builder.InsertChart(ChartType.Pie, 432, 252);
            var chart = shape.Chart;
            chart.Title.Text = "Produced by Aspose.Words Processor plugin.";

            chart.Series.Clear();
            chart.Series.Add("Series 1",
                new string[] { "Category 1", "Category 2", "Category 3" },
                new double[] { 2.7, 3.2, 0.8 });

            doc.Save(ArtifactsDir + "ProcessorWordPlugin.CreateChartWord.docx");
            //ExEnd:CreateChartWord
        }

        [Test]
        public void CreateBookmarkWord()
        {
            //ExStart:CreateBookmarkWord
            //GistId:c3b9534ddda2dbbd6f267b53dda5b605
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.StartBookmark("Bookmark");
            builder.Write("Produced by Aspose.Words Processor plugin.");
            builder.EndBookmark("Bookmark");

            doc.Save(ArtifactsDir + "ProcessorWordPlugin.CreateBookmarkWord.docx");
            //ExEnd:CreateBookmarkWord
        }

        [Test]
        public void EditDocumentWord()
        {
            //ExStart:EditDocumentWord
            //GistId:c3b9534ddda2dbbd6f267b53dda5b605
            var doc = new Document(MyDir + "Document.docx");
            var builder = new DocumentBuilder(doc);

            builder.MoveToDocumentEnd();
            builder.Writeln("Produced by Aspose.Words Processor plugin.");

            doc.Save(ArtifactsDir + "ProcessorWordPlugin.EditDocumentWord.docx");
            //ExEnd:EditDocumentWord
        }
    }
}
