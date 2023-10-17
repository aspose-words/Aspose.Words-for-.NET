using Aspose.Words.Drawing.Charts;
using Aspose.Words;
using NUnit.Framework;
using Aspose.Words.Saving;

namespace PluginsExamples
{
    public class ProcessorXmlPlugin : PluginsExamplesBase
    {
        [Test]
        public void CreateChartXml()
        {
            //ExStart:CreateChartXml
            //GistId:c3d0fd3f2d557c95863ed7f29b0ab66b
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            var shape = builder.InsertChart(ChartType.Pie, 432, 252);
            var chart = shape.Chart;
            chart.Title.Text = "Produced by Aspose.Words Processor plugin.";

            chart.Series.Clear();
            chart.Series.Add("Series 1",
                new string[] { "Category 1", "Category 2", "Category 3" },
                new double[] { 2.7, 3.2, 0.8 });

            doc.Save(ArtifactsDir + "ProcessorXmlPlugin.CreateChartXml.xps");
            //ExEnd:CreateChartXml
        }

        [Test]
        public void CreateBookmarkXml()
        {
            //ExStart:CreateBookmarkXml
            //GistId:c3d0fd3f2d557c95863ed7f29b0ab66b
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.StartBookmark("Bookmark");
            builder.Write("Produced by Aspose.Words Processor plugin.");
            builder.EndBookmark("Bookmark");

            var saveOptions = new XpsSaveOptions();
            saveOptions.OutlineOptions.BookmarksOutlineLevels.Add("Bookmark", 1);

            doc.Save(ArtifactsDir + "ProcessorXmlPlugin.CreateBookmarkXml.xps", saveOptions);
            //ExEnd:CreateBookmarkXml
        }

        [Test]
        public void EditDocumentXml()
        {
            //ExStart:EditDocumentXml
            //GistId:c3d0fd3f2d557c95863ed7f29b0ab66b
            var doc = new Document(MyDir + "Document.xml");
            var builder = new DocumentBuilder(doc);

            builder.MoveToDocumentEnd();
            builder.Writeln("Produced by Aspose.Words Processor plugin.");

            doc.Save(ArtifactsDir + "ProcessorXmlPlugin.EditDocumentXml.xps");
            //ExEnd:EditDocumentXml
        }
    }
}
