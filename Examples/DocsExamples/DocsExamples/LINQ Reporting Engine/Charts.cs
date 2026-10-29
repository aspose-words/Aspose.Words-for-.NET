using Aspose.Words;
using Aspose.Words.Reporting;
using DocsExamples.LINQ_Reporting_Engine.Helpers;
using NUnit.Framework;

namespace DocsExamples.LINQ_Reporting_Engine
{
    public class Charts : DocsExamplesBase
    {
        [Test]
        public void CreateBubbleChart()
        {
            //ExStart:BubbleChart
            Document doc = new Document(MyDir + "Reporting engine template - Bubble chart.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Common.GetManagers(), "managers");
            
            doc.Save(ArtifactsDir + "ReportingEngine.CreateBubbleChart.docx");
            //ExEnd:BubbleChart
        }

        [Test]
        public void ChartSeriesNameDynamically()
        {
            //ExStart:ChartSeriesNameDynamically
            //GistId:396ec359e1983055f29e969c4497dc7b
            Document doc = new Document(MyDir + "Reporting engine template - Chart.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Common.GetManagers(), "managers");

            doc.Save(ArtifactsDir + "ReportingEngine.ChartSeriesNameDynamically.docx");
            //ExEnd:ChartSeriesNameDynamically
        }

        [Test]
        public void ChartWithFilteringGroupingOrdering()
        {
            //ExStart:ChartWithFilteringGroupingOrdering
            Document doc = new Document(MyDir + "Reporting engine template - Chart with filtering.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetContracts(), "contracts");

            doc.Save(ArtifactsDir + "ReportingEngine.ChartWithFilteringGroupingOrdering.docx");
            //ExEnd:ChartWithFilteringGroupingOrdering
        }

        [Test]
        public void PieChart()
        {
            //ExStart:PieChart
            Document doc = new Document(MyDir + "Reporting engine template - Pie chart.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetManagers(), "managers");

            doc.Save(ArtifactsDir + "ReportingEngine.PieChart.docx");
            //ExEnd:PieChart
        }

        [Test]
        public void ScatterChart()
        {
            //ExStart:ScatterChart
            //GistId:688c5a950ba6b1a9c17114eb1260f4c9
            Document doc = new Document(MyDir + "Reporting engine template - Scatter chart.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, Helpers.Common.GetContracts(), "contracts");

            doc.Save(ArtifactsDir + "ReportingEngine.ScatterChart.docx");
            //ExEnd:ScatterChart
        }
    }
}