// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Reporting;
using NUnit.Framework;
using System;
using System.IO;

namespace ApiExamples.LinqTemp
{
    [TestFixture]
    public class ExReportingEngine : ApiExampleBase
    {
        [Test]
        public void BuildingColumnChart()
        {
            //ExStart:BuildingColumnChart
            //GistId:a9bfce4e06620c7bb2f1f0af6d166f0e
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Column Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Column Chart Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Column Chart Report.docx");
            //ExEnd:BuildingColumnChart

            // Test the report.
            CompareDocs("Column Chart Report.docx", "Column Chart Report Gold.docx");
        }

        [Test]
        public void BuildingLineChart()
        {
            //ExStart:BuildingLineChart
            //GistId:2dcdbb630b51b99c6bd65a6603dfd982
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Line Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Line Chart Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Line Chart Report.docx");
            //ExEnd:BuildingLineChart

            // Test the report.
            CompareDocs("Line Chart Report.docx", "Line Chart Report Gold.docx");
        }

        [Test]
        public void BuildingPieChart()
        {
            //ExStart:BuildingPieChart
            //GistId:a42494e6d04efdee2f0d39807dadb5d6
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Pie Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Pie Chart Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Pie Chart Report.docx");
            //ExEnd:BuildingPieChart

            // Test the report.
            CompareDocs("Pie Chart Report.docx", "Pie Chart Report Gold.docx");
        }

        [Test]
        public void BuildingDonutChart()
        {
            //ExStart:BuildingDonutChart
            //GistId:d34c252777a718d0a7c2867f71bf0888
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Donut Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Donut Chart Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Donut Chart Report.docx");
            //ExEnd:BuildingDonutChart

            // Test the report.
            CompareDocs("Donut Chart Report.docx", "Donut Chart Report Gold.docx");
        }

        [Test]
        public void BuildingBarChart()
        {
            //ExStart:BuildingBarChart
            //GistId:b28dd058a87a3f9ce67a797e9132bdf6
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Bar Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Bar Chart Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Bar Chart Report.docx");
            //ExEnd:BuildingBarChart

            // Test the report.
            CompareDocs("Bar Chart Report.docx", "Bar Chart Report Gold.docx");
        }

        [Test]
        public void BuildingAreaChart()
        {
            //ExStart:BuildingAreaChart
            //GistId:bc0e095ec0072a1e1e704777334b5023
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Area Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Area Chart Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Area Chart Report.docx");
            //ExEnd:BuildingAreaChart

            // Test the report.
            CompareDocs("Area Chart Report.docx", "Area Chart Report Gold.docx");
        }

        [Test]
        public void BuildingScatterChart()
        {
            //ExStart:BuildingScatterChart
            //GistId:afe1b02e007250935be30b9456295e21
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Scatter Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Scatter Chart Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Scatter Chart Report.docx");
            //ExEnd:BuildingScatterChart

            // Test the report.
            CompareDocs("Scatter Chart Report.docx", "Scatter Chart Report Gold.docx");
        }

        [Test]
        public void BuildingBubbleChart()
        {
            //ExStart:BuildingBubbleChart
            //GistId:595f34976fce9c39136e0a8ccf5241fd
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Bubble Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Bubble Chart Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Bubble Chart Report.docx");
            //ExEnd:BuildingBubbleChart

            // Test the report.
            CompareDocs("Bubble Chart Report.docx", "Bubble Chart Report Gold.docx");
        }

        [Test]
        public void BuildingSurfaceChart()
        {
            //ExStart:BuildingSurfaceChart
            //GistId:2fcc8b5e3ad0fd49a3528563331cbe66
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Surface Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Surface Chart Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Surface Chart Report.docx");
            //ExEnd:BuildingSurfaceChart

            // Test the report.
            CompareDocs("Surface Chart Report.docx", "Surface Chart Report Gold.docx");
        }

        [Test]
        public void BuildingRadarChart()
        {
            //ExStart:BuildingRadarChart
            //GistId:96bd2aea26f54c0ddfd32a6301cfd32e
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Radar Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Radar Chart Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Radar Chart Report.docx");
            //ExEnd:BuildingRadarChart

            // Test the report.
            CompareDocs("Radar Chart Report.docx", "Radar Chart Report Gold.docx");
        }

        [Test]
        public void BuildingTreemapChart()
        {
            //ExStart:BuildingTreemapChart
            //GistId:0564a1c00008d6da73c2a1fc9c744794
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Treemap Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Treemap Chart Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Treemap Chart Report.docx");
            //ExEnd:BuildingTreemapChart

            // Test the report.
            CompareDocs("Treemap Chart Report.docx", "Treemap Chart Report Gold.docx");
        }

        [Test]
        public void BuildingSunburstChart()
        {
            //ExStart:BuildingSunburstChart
            //GistId:ae0d2a688336922adc838a9a7225d1cc
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Sunburst Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Sunburst Chart Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Sunburst Chart Report.docx");
            //ExEnd:BuildingSunburstChart

            // Test the report.
            CompareDocs("Sunburst Chart Report.docx", "Sunburst Chart Report Gold.docx");
        }

        [Test]
        public void BuildingHistogramChart()
        {
            //ExStart:BuildingHistogramChart
            //GistId:edfb65be6c8467ae3fef7574f138218f
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Histogram Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Histogram Chart Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Histogram Chart Report.docx");
            //ExEnd:BuildingHistogramChart

            // Test the report.
            CompareDocs("Histogram Chart Report.docx", "Histogram Chart Report Gold.docx");
        }

        [Test]
        public void BuildingParetoChart()
        {
            //ExStart:BuildingParetoChart
            //GistId:45210245bf418e3146e39dd905818695
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Pareto Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Pareto Chart Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Pareto Chart Report.docx");
            //ExEnd:BuildingParetoChart

            // Test the report.
            CompareDocs("Pareto Chart Report.docx", "Pareto Chart Report Gold.docx");
        }

        [Test]
        public void BuildingBoxAndWhiskerChart()
        {
            //ExStart:BuildingBoxAndWhiskerChart
            //GistId:a0b8163b5ecc129f8c439cc0b44a85d3
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Box and Whisker Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Box and Whisker Chart Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Box and Whisker Chart Report.docx");
            //ExEnd:BuildingBoxAndWhiskerChart

            // Test the report.
            CompareDocs("Box and Whisker Chart Report.docx", "Box and Whisker Chart Report Gold.docx");
        }

        [Test]
        public void BuildingWaterfallChart()
        {
            //ExStart:BuildingWaterfallChart
            //GistId:3100394ea4660bad51382db4da276d4a
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Waterfall Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Waterfall Chart Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Waterfall Chart Report.docx");
            //ExEnd:BuildingWaterfallChart

            // Test the report.
            CompareDocs("Waterfall Chart Report.docx", "Waterfall Chart Report Gold.docx");
        }

        [Test]
        public void BuildingFunnelChart()
        {
            //ExStart:BuildingFunnelChart
            //GistId:407b901cd82f93eabd84bdc6dd2cb5b7
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Funnel Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Funnel Chart Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Funnel Chart Report.docx");
            //ExEnd:BuildingFunnelChart

            // Test the report.
            CompareDocs("Funnel Chart Report.docx", "Funnel Chart Report Gold.docx");
        }

        [Test]
        public void BuildingStockChart()
        {
            //ExStart:BuildingStockChart
            //GistId:bbcd237a7d0afc8f48f2e9d39ced30e7
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Stock Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Stock Chart Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Stock Chart Report.docx");
            //ExEnd:BuildingStockChart

            // Test the report.
            CompareDocs("Stock Chart Report.docx", "Stock Chart Report Gold.docx");
        }

        [Test]
        public void BuildingComboChart()
        {
            //ExStart:BuildingComboChart
            //GistId:ce435eace52f0504985b9f10e7203e3d
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Combo Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Combo Chart Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Combo Chart Report.docx");
            //ExEnd:BuildingComboChart

            // Test the report.
            CompareDocs("Combo Chart Report.docx", "Combo Chart Report Gold.docx");
        }

        /// <summary>
        /// A helper method asserting that two files are equal.
        /// </summary>
        /// <param name="artifactFileName">The name of a file within <see cref="ArtifactsDir"/>.</param>
        /// <param name="linqGoldFileName">The name of a file within <see cref="LinqGoldsDir"/>.</param>
        private static void CompareDocs(string artifactFileName, string linqGoldFileName)
        {
            string artifactPath = ArtifactsDir + artifactFileName;
            string linqGoldPath = LinqGoldsDir + linqGoldFileName;

            if (!File.Exists(linqGoldPath))
                File.Copy(artifactPath, linqGoldPath);
            else
                Assert.IsTrue(DocumentHelper.CompareDocs(artifactPath, linqGoldPath));
        }

        /// <summary>
        /// Gets the path to the directory with template and data files used by this class. Ends with a back slash.
        /// </summary>
        private static string MyLinqDir { get; }

        /// <summary>
        /// Gets the path to the directory with previously generated reports used by this class. Ends with a back slash.
        /// </summary>
        private static string LinqGoldsDir { get; }

        static ExReportingEngine()
        {
            MyLinqDir = new Uri(new Uri(MyDir), @"LINQ/").LocalPath;
            LinqGoldsDir = new Uri(new Uri(GoldsDir), @"LINQ/").LocalPath;
        }
    }
}
