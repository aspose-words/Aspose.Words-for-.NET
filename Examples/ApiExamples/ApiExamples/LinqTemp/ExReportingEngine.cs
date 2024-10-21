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

        [Test]
        public void BuildingGaugeChart()
        {
            //ExStart:BuildingGaugeChart
            //GistId:45e82d351f10b65ceca440d95d72aa5a
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Gauge Chart Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Gauge Chart Data.json");

            // Build a report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource);

            // Save the report.
            doc.Save(ArtifactsDir + "Gauge Chart Report.docx");
            //ExEnd:BuildingGaugeChart

            // Test the report.
            CompareDocs("Gauge Chart Report.docx", "Gauge Chart Report Gold.docx");
        }

        [Test]
        public void BuildingChartWithVariableNumberOfSeries1()
        {
            //ExStart:BuildingChartWithVariableNumberOfSeries
            //GistId:2f7a15f32189afa4c8a2195926cedf7d
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Chart with Variable Number of Series Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Chart with Variable Number of Series Data 1.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "data");

            // Save the report.
            doc.Save(ArtifactsDir + "Chart with Variable Number of Series Report 1.docx");
            //ExEnd:BuildingChartWithVariableNumberOfSeries

            // Test the report.
            CompareDocs("Chart with Variable Number of Series Report 1.docx",
                "Chart with Variable Number of Series Report 1 Gold.docx");
        }

        [Test]
        public void BuildingChartWithVariableNumberOfSeries2()
        {
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Chart with Variable Number of Series Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Chart with Variable Number of Series Data 2.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "data");

            // Save the report.
            doc.Save(ArtifactsDir + "Chart with Variable Number of Series Report 2.docx");
            
            // Test the report.
            CompareDocs("Chart with Variable Number of Series Report 2.docx",
                "Chart with Variable Number of Series Report 2 Gold.docx");
        }

        [Test]
        public void ChangingChartTitle()
        {
            //ExStart:ChangingChartTitle
            //GistId:86abd6970b9db97fa6b8a1a5852ac2fc
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Chart with Changing Title Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Chart with Changing Title Data.json");

            // Build a report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource);

            // Save the report.
            doc.Save(ArtifactsDir + "Chart with Changing Title Report.docx");
            //ExEnd:ChangingChartTitle

            // Test the report.
            CompareDocs("Chart with Changing Title Report.docx", "Chart with Changing Title Report Gold.docx");
        }

        [Test]
        public void ChangingChartLegendEntry()
        {
            //ExStart:ChangingChartLegendEntry
            //GistId:ad4fcc877c14e6d4bd7346fce06d027a
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Chart with Changing Legend Entries Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Chart with Changing Legend Entries Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "data");

            // Save the report.
            doc.Save(ArtifactsDir + "Chart with Changing Legend Entries Report.docx");
            //ExEnd:ChangingChartLegendEntry

            // Test the report.
            CompareDocs("Chart with Changing Legend Entries Report.docx",
                "Chart with Changing Legend Entries Report Gold.docx");
        }

        [Test]
        public void ChangingChartAxisTitle()
        {
            //ExStart:ChangingChartAxisTitle
            //GistId:e9668c5d7f03033c56bea98921f80e72
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Chart with Changing Axis Titles Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Chart with Changing Axis Titles Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "data");

            // Save the report.
            doc.Save(ArtifactsDir + "Chart with Changing Axis Titles Report.docx");
            //ExEnd:ChangingChartAxisTitle

            // Test the report.
            CompareDocs("Chart with Changing Axis Titles Report.docx", "Chart with Changing Axis Titles Report Gold.docx");
        }

        [Test]
        public void ChangingChartSeriesColor()
        {
            //ExStart:ChangingChartSeriesColor
            //GistId:f9a631509dab79e8cc36f2905c8e17d2
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Chart with Changing Series Colors Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Chart with Changing Series Colors Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "data");

            // Save the report.
            doc.Save(ArtifactsDir + "Chart with Changing Series Colors Report.docx");
            //ExEnd:ChangingChartSeriesColor

            // Test the report.
            CompareDocs("Chart with Changing Series Colors Report.docx", "Chart with Changing Series Colors Report Gold.docx");
        }

        [Test]
        public void ChangingChartSeriesPointColor()
        {
            //ExStart:ChangingChartSeriesPointColor
            //GistId:efa74853ed5a592fa775479b0b480651
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Chart with Changing Series Point Colors Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Chart with Changing Series Point Colors Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Chart with Changing Series Point Colors Report.docx");
            //ExEnd:ChangingChartSeriesPointColor

            // Test the report.
            CompareDocs("Chart with Changing Series Point Colors Report.docx",
                "Chart with Changing Series Point Colors Report Gold.docx");
        }

        [Test]
        public void BuildingTable()
        {
            //ExStart:BuildingTable
            //GistId:1db755b118593b067e1de46ad5fbe550
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Table Data.json");

            // Build a report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource);

            // Save the report.
            doc.Save(ArtifactsDir + "Table Report.docx");
            //ExEnd:BuildingTable

            // Test the report.
            CompareDocs("Table Report.docx", "Table Report Gold.docx");
        }

        [Test]
        public void BindingTableRowsToCollection()
        {
            //ExStart:BindingTableRowsToCollection
            //GistId:046e728500b12abd8ed5c68274657d55
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Rows Bound to Collection Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Table with Rows Bound to Collection Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Rows Bound to Collection Report.docx");
            //ExEnd:BindingTableRowsToCollection

            // Test the report.
            CompareDocs("Table with Rows Bound to Collection Report.docx",
                "Table with Rows Bound to Collection Report Gold.docx");
        }

        [Test]
        public void ChangingTableHeaders()
        {
            //ExStart:ChangingTableHeaders
            //GistId:16ccc82dfb99f3489159a9462aa3daab
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Changing Headers Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Table with Changing Headers Data.json");

            // Build a report.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource);

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Changing Headers Report.docx");
            //ExEnd:ChangingTableHeaders

            // Test the report.
            CompareDocs("Table with Changing Headers Report.docx",
                "Table with Changing Headers Report Gold.docx");
        }

        [Test]
        public void AddingTotalToTable()
        {
            //ExStart:AddingTotalToTable
            //GistId:a7c74dbddf71e28c54b1ae984081a37e
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Total Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Table with Total Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Total Report.docx");
            //ExEnd:AddingTotalToTable

            // Test the report.
            CompareDocs("Table with Total Report.docx", "Table with Total Report Gold.docx");
        }

        [Test]
        public void DisplayingMessageForEmptyTable1()
        {
            //ExStart:DisplayingMessageForEmptyTable
            //GistId:42c09d5494a9a9ba8bfe8ec5a13d8fff
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Message If Empty Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Table with Message If Empty Data 1.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.Options |= ReportBuildOptions.AllowMissingMembers; // Needed to accept possibly missing parts of data.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Message If Empty Report 1.docx");
            //ExEnd:DisplayingMessageForEmptyTable

            // Test the report.
            CompareDocs("Table with Message If Empty Report 1.docx",
                "Table with Message If Empty Report 1 Gold.docx");
        }

        [Test]
        public void DisplayingMessageForEmptyTable2()
        {
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Message If Empty Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Table with Message If Empty Data 2.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.Options |= ReportBuildOptions.AllowMissingMembers; // Needed to accept possibly missing parts of data.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Message If Empty Report 2.docx");

            // Test the report.
            CompareDocs("Table with Message If Empty Report 2.docx",
                "Table with Message If Empty Report 2 Gold.docx");
        }

        [Test]
        public void ShowingTableRowBasedOnCondition1()
        {
            //ExStart:ShowingTableRowBasedOnCondition
            //GistId:d8afeb97e4bbc5a447e225cd8a7da4cb
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Row Shown Based on Condition Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Table with Row Shown Based on Condition Data 1.json");

            // Build a report.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource);

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Row Shown Based on Condition Report 1.docx");
            //ExEnd:ShowingTableRowBasedOnCondition

            // Test the report.
            CompareDocs("Table with Row Shown Based on Condition Report 1.docx",
                "Table with Row Shown Based on Condition Report 1 Gold.docx");
        }

        [Test]
        public void ShowingTableRowBasedOnCondition2()
        {
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Row Shown Based on Condition Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Table with Row Shown Based on Condition Data 2.json");

            // Build a report.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource);

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Row Shown Based on Condition Report 2.docx");

            // Test the report.
            CompareDocs("Table with Row Shown Based on Condition Report 2.docx",
                "Table with Row Shown Based on Condition Report 2 Gold.docx");
        }

        [Test]
        public void BindingTableRowsToCollectionBasedOnCondition()
        {
            //ExStart:BindingTableRowsToCollectionBasedOnCondition
            //GistId:43148c116d8d92bfb9c1286ed68ddf23
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Rows Bound to Collection Based on Condition Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(
                MyLinqDir + "Table with Rows Bound to Collection Based on Condition Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Rows Bound to Collection Based on Condition Report.docx");
            //ExEnd:BindingTableRowsToCollectionBasedOnCondition

            // Test the report.
            CompareDocs("Table with Rows Bound to Collection Based on Condition Report.docx",
                "Table with Rows Bound to Collection Based on Condition Report Gold.docx");
        }

        [Test]
        public void ApplyingConditionalFormattingToTableRows()
        {
            //ExStart:ApplyingConditionalFormattingToTableRows
            //GistId:749070be98dfa2c3159257b4f3ea3401
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Conditional Formatting Applied to Rows Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(
                MyLinqDir + "Table with Conditional Formatting Applied to Rows Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Conditional Formatting Applied to Rows Report.docx");
            //ExEnd:ApplyingConditionalFormattingToTableRows

            // Test the report.
            CompareDocs("Table with Conditional Formatting Applied to Rows Report.docx",
                "Table with Conditional Formatting Applied to Rows Report Gold.docx");
        }

        [Test]
        public void ApplyingBackgroundColorsToTableRows()
        {
            //ExStart:ApplyingBackgroundColorsToTableRows
            //GistId:b1c9ffb0ed27ad32a33f4078b0b75231
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Background Colors Applied to Rows Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(
                MyLinqDir + "Table with Background Colors Applied to Rows Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Background Colors Applied to Rows Report.docx");
            //ExEnd:ApplyingBackgroundColorsToTableRows

            // Test the report.
            CompareDocs("Table with Background Colors Applied to Rows Report.docx",
                "Table with Background Colors Applied to Rows Report Gold.docx");
        }

        [Test]
        public void ApplyingTextColorsToTableRows()
        {
            //ExStart:ApplyingTextColorsToTableRows
            //GistId:2b490da2afa4b725e1b0ceefaec92b32
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Text Colors Applied to Rows Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(
                MyLinqDir + "Table with Text Colors Applied to Rows Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Text Colors Applied to Rows Report.docx");
            //ExEnd:ApplyingTextColorsToTableRows

            // Test the report.
            CompareDocs("Table with Text Colors Applied to Rows Report.docx",
                "Table with Text Colors Applied to Rows Report Gold.docx");
        }

        [Test]
        public void AddingRunningTotalToTable()
        {
            //ExStart:AddingRunningTotalToTable
            //GistId:c7c5f2b0a30d1c83315c48a3c47639ec
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Running Total Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Table with Running Total Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Running Total Report.docx");
            //ExEnd:AddingRunningTotalToTable

            // Test the report.
            CompareDocs("Table with Running Total Report.docx", "Table with Running Total Report Gold.docx");
        }

        [Test]
        public void DisplayingSeveralItemsPerTableRow()
        {
            //ExStart:DisplayingSeveralItemsPerTableRow
            //GistId:6d983652a385c662c3a60602d4e54f16
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Several Items Displayed per Row Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Table with Several Items Displayed per Row Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Several Items Displayed per Row Report.docx");
            //ExEnd:DisplayingSeveralItemsPerTableRow

            // Test the report.
            CompareDocs("Table with Several Items Displayed per Row Report.docx",
                "Table with Several Items Displayed per Row Report Gold.docx");
        }

        [Test]
        public void DisplayingMasterDetailDataInOneTableOption1()
        {
            //ExStart:DisplayingMasterDetailDataInOneTableOption1
            //GistId:e372647a722c7b3c2a0c8abd1d8fe572
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Master-Detail Data Option 1 Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Table with Master-Detail Data Option 1 Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Master-Detail Data Option 1 Report.docx");
            //ExEnd:DisplayingMasterDetailDataInOneTableOption1

            // Test the report.
            CompareDocs("Table with Master-Detail Data Option 1 Report.docx",
                "Table with Master-Detail Data Option 1 Report Gold.docx");
        }

        [Test]
        public void DisplayingMasterDetailDataInOneTableOption2()
        {
            //ExStart:DisplayingMasterDetailDataInOneTableOption2
            //GistId:30e09bd0914ebed820339a602432e101
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Master-Detail Data Option 2 Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Table with Master-Detail Data Option 2 Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Master-Detail Data Option 2 Report.docx");
            //ExEnd:DisplayingMasterDetailDataInOneTableOption2

            // Test the report.
            CompareDocs("Table with Master-Detail Data Option 2 Report.docx",
                "Table with Master-Detail Data Option 2 Report Gold.docx");
        }

        [Test]
        public void AddingSubheaderToTable()
        {
            //ExStart:AddingSubheaderToTable
            //GistId:f1dae907de53577beb4989c7ae6b2cc8
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Subheaders Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Table with Subheaders Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Subheaders Report.docx");
            //ExEnd:AddingSubheaderToTable

            // Test the report.
            CompareDocs("Table with Subheaders Report.docx", "Table with Subheaders Report Gold.docx");
        }

        [Test]
        public void AddingSubtotalToTable()
        {
            //ExStart:AddingSubtotalToTable
            //GistId:b84b813cf9c0a0e91b1fad7e183ffc2f
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Subtotals Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Table with Subtotals Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Subtotals Report.docx");
            //ExEnd:AddingSubtotalToTable

            // Test the report.
            CompareDocs("Table with Subtotals Report.docx", "Table with Subtotals Report Gold.docx");
        }

        [Test]
        public void BuildingSingleColumnTable()
        {
            //ExStart:BuildingSingleColumnTable
            //GistId:da447f6a823cc5e5c0599122edba80cc
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Single-Column Table Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Single-Column Table Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Single-Column Table Report.docx");
            //ExEnd:BuildingSingleColumnTable

            // Test the report.
            CompareDocs("Single-Column Table Report.docx", "Single-Column Table Report Gold.docx");
        }

        [Test]
        public void ShowingSingleColumnTableRowBasedOnCondition1()
        {
            //ExStart:ShowingSingleColumnTableRowBasedOnCondition
            //GistId:b841d4c42cfcad6ede6feaaf84f485ad
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Single-Column Table with Row Shown Based on Condition Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(
                MyLinqDir + "Single-Column Table with Row Shown Based on Condition Data 1.json");

            // Build a report.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource);

            // Save the report.
            doc.Save(ArtifactsDir + "Single-Column Table with Row Shown Based on Condition Report 1.docx");
            //ExEnd:ShowingSingleColumnTableRowBasedOnCondition

            // Test the report.
            CompareDocs("Single-Column Table with Row Shown Based on Condition Report 1.docx",
                "Single-Column Table with Row Shown Based on Condition Report 1 Gold.docx");
        }

        [Test]
        public void ShowingSingleColumnTableRowBasedOnCondition2()
        {
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Single-Column Table with Row Shown Based on Condition Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(
                MyLinqDir + "Single-Column Table with Row Shown Based on Condition Data 2.json");

            // Build a report.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource);

            // Save the report.
            doc.Save(ArtifactsDir + "Single-Column Table with Row Shown Based on Condition Report 2.docx");

            // Test the report.
            CompareDocs("Single-Column Table with Row Shown Based on Condition Report 2.docx",
                "Single-Column Table with Row Shown Based on Condition Report 2 Gold.docx");
        }

        [Test]
        public void BindingSingleColumnTableRowsToCollectionBasedOnCondition()
        {
            //ExStart:BindingSingleColumnTableRowsToCollectionBasedOnCondition
            //GistId:091dc5121b1d33ca66031aa3c04b72f9
            // Open the template document.
            Document doc = new Document(
                MyLinqDir + "Single-Column Table with Rows Bound to Collection Based on Condition Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(
                MyLinqDir + "Single-Column Table with Rows Bound to Collection Based on Condition Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Single-Column Table with Rows Bound to Collection Based on Condition Report.docx");
            //ExEnd:BindingSingleColumnTableRowsToCollectionBasedOnCondition

            // Test the report.
            CompareDocs("Single-Column Table with Rows Bound to Collection Based on Condition Report.docx",
                "Single-Column Table with Rows Bound to Collection Based on Condition Report Gold.docx");
        }

        [Test]
        public void BuildingHorizontalTable()
        {
            //ExStart:BuildingHorizontalTable
            //GistId:babe28211ce3b225949a73ada8989ead
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Horizontal Table Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Horizontal Table Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Horizontal Table Report.docx");
            //ExEnd:BuildingHorizontalTable

            // Test the report.
            CompareDocs("Horizontal Table Report.docx", "Horizontal Table Report Gold.docx");
        }

        [Test]
        public void AddingTotalToHorizontalTable()
        {
            //ExStart:AddingTotalToHorizontalTable
            //GistId:402615eda5e0a727e955e979d7b8790b
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Horizontal Table with Total Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Horizontal Table with Total Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Horizontal Table with Total Report.docx");
            //ExEnd:AddingTotalToHorizontalTable

            // Test the report.
            CompareDocs("Horizontal Table with Total Report.docx", "Horizontal Table with Total Report Gold.docx");
        }

        [Test]
        public void DisplayingMessageForEmptyHorizontalTable1()
        {
            //ExStart:DisplayingMessageForEmptyHorizontalTable
            //GistId:5bcd601ff2fbed17945d12700be84194
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Horizontal Table with Message If Empty Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Horizontal Table with Message If Empty Data 1.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.Options |= ReportBuildOptions.AllowMissingMembers; // Needed to accept possibly missing parts of data.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Horizontal Table with Message If Empty Report 1.docx");
            //ExEnd:DisplayingMessageForEmptyHorizontalTable

            // Test the report.
            CompareDocs("Horizontal Table with Message If Empty Report 1.docx",
                "Horizontal Table with Message If Empty Report 1 Gold.docx");
        }

        [Test]
        public void DisplayingMessageForEmptyHorizontalTable2()
        {
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Horizontal Table with Message If Empty Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Horizontal Table with Message If Empty Data 2.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.Options |= ReportBuildOptions.AllowMissingMembers; // Needed to accept possibly missing parts of data.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Horizontal Table with Message If Empty Report 2.docx");

            // Test the report.
            CompareDocs("Horizontal Table with Message If Empty Report 2.docx",
                "Horizontal Table with Message If Empty Report 2 Gold.docx");
        }

        [Test]
        public void DisplayingMasterDetailDataInOneHorizontalTableOption1()
        {
            //ExStart:DisplayingMasterDetailDataInOneHorizontalTableOption1
            //GistId:28c0c52d19a2a5ea35b9ae6bae0f9b36
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Horizontal Table with Master-Detail Data Option 1 Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(
                MyLinqDir + "Horizontal Table with Master-Detail Data Option 1 Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Horizontal Table with Master-Detail Data Option 1 Report.docx");
            //ExEnd:DisplayingMasterDetailDataInOneHorizontalTableOption1

            // Test the report.
            CompareDocs("Horizontal Table with Master-Detail Data Option 1 Report.docx",
                "Horizontal Table with Master-Detail Data Option 1 Report Gold.docx");
        }

        [Test]
        public void DisplayingMasterDetailDataInOneHorizontalTableOption2()
        {
            //ExStart:DisplayingMasterDetailDataInOneHorizontalTableOption2
            //GistId:5626903685eaf2c1a5c640977da5f7e8
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Horizontal Table with Master-Detail Data Option 2 Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(
                MyLinqDir + "Horizontal Table with Master-Detail Data Option 2 Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Horizontal Table with Master-Detail Data Option 2 Report.docx");
            //ExEnd:DisplayingMasterDetailDataInOneHorizontalTableOption2

            // Test the report.
            CompareDocs("Horizontal Table with Master-Detail Data Option 2 Report.docx",
                "Horizontal Table with Master-Detail Data Option 2 Report Gold.docx");
        }

        [Test]
        public void ShowingTableColumnBasedOnCondition1()
        {
            //ExStart:ShowingTableColumnBasedOnCondition
            //GistId:f3ef3b2284ff6716aae500ba6d82b68e
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Column Shown Based on Condition Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(
                MyLinqDir + "Table with Column Shown Based on Condition Data 1.json");

            // Build a report.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource);

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Column Shown Based on Condition Report 1.docx");
            //ExEnd:ShowingTableColumnBasedOnCondition

            // Test the report.
            CompareDocs("Table with Column Shown Based on Condition Report 1.docx",
                "Table with Column Shown Based on Condition Report 1 Gold.docx");
        }

        [Test]
        public void ShowingTableColumnBasedOnCondition2()
        {
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Column Shown Based on Condition Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(
                MyLinqDir + "Table with Column Shown Based on Condition Data 2.json");

            // Build a report.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource);

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Column Shown Based on Condition Report 2.docx");

            // Test the report.
            CompareDocs("Table with Column Shown Based on Condition Report 2.docx",
                "Table with Column Shown Based on Condition Report 2 Gold.docx");
        }

        [Test]
        public void ShowingColumnOfTableWithRowsBoundToCollectionBasedOnCondition1()
        {
            //ExStart:ShowingColumnOfTableWithRowsBoundToCollectionBasedOnCondition
            //GistId:25cc19f51209ad36cccab1d8bb4f4475
            // Open the template document.
            Document doc = new Document(
                MyLinqDir + "Table with Column Shown Based on Condition and Rows Bound to Collection Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(
                MyLinqDir + "Table with Column Shown Based on Condition and Rows Bound to Collection Data 1.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "ds");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Column Shown Based on Condition and Rows Bound to Collection Report 1.docx");
            //ExEnd:ShowingColumnOfTableWithRowsBoundToCollectionBasedOnCondition

            // Test the report.
            CompareDocs("Table with Column Shown Based on Condition and Rows Bound to Collection Report 1.docx",
                "Table with Column Shown Based on Condition and Rows Bound to Collection Report 1 Gold.docx");
        }

        [Test]
        public void ShowingColumnOfTableWithRowsBoundToCollectionBasedOnCondition2()
        {
            // Open the template document.
            Document doc = new Document(
                MyLinqDir + "Table with Column Shown Based on Condition and Rows Bound to Collection Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(
                MyLinqDir + "Table with Column Shown Based on Condition and Rows Bound to Collection Data 2.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "ds");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Column Shown Based on Condition and Rows Bound to Collection Report 2.docx");
            
            // Test the report.
            CompareDocs("Table with Column Shown Based on Condition and Rows Bound to Collection Report 2.docx",
                "Table with Column Shown Based on Condition and Rows Bound to Collection Report 2 Gold.docx");
        }

        [Test]
        public void BindingTableColumnsToCollectionBasedOnCondition()
        {
            //ExStart:BindingTableColumnsToCollectionBasedOnCondition
            //GistId:d30c7e45035ae070be2e511f87fce5db
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Columns Bound to Collection Based on Condition Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(
                MyLinqDir + "Table with Columns Bound to Collection Based on Condition Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Columns Bound to Collection Based on Condition Report.docx");
            //ExEnd:BindingTableColumnsToCollectionBasedOnCondition

            // Test the report.
            CompareDocs("Table with Columns Bound to Collection Based on Condition Report.docx",
                "Table with Columns Bound to Collection Based on Condition Report Gold.docx");
        }

        [Test]
        public void ApplyingConditionalFormattingToTableColumns()
        {
            //ExStart:ApplyingConditionalFormattingToTableColumns
            //GistId:896a6e0853a7e706952eec3203954f4b
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Conditional Formatting Applied to Columns Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(
                MyLinqDir + "Table with Conditional Formatting Applied to Columns Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Conditional Formatting Applied to Columns Report.docx");
            //ExEnd:ApplyingConditionalFormattingToTableColumns

            // Test the report.
            CompareDocs("Table with Conditional Formatting Applied to Columns Report.docx",
                "Table with Conditional Formatting Applied to Columns Report Gold.docx");
        }

        [Test]
        public void ApplyingBackgroundColorsToTableColumns()
        {
            //ExStart:ApplyingBackgroundColorsToTableColumns
            //GistId:aae874038596ea2e96371bb37b30b692
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Background Colors Applied to Columns Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(
                MyLinqDir + "Table with Background Colors Applied to Columns Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Background Colors Applied to Columns Report.docx");
            //ExEnd:ApplyingBackgroundColorsToTableColumns

            // Test the report.
            CompareDocs("Table with Background Colors Applied to Columns Report.docx",
                "Table with Background Colors Applied to Columns Report Gold.docx");
        }

        [Test]
        public void ApplyingTextColorsToTableColumns()
        {
            //ExStart:ApplyingTextColorsToTableColumns
            //GistId:ad728469bf0d8b577d99b4d57004b59e
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Text Colors Applied to Columns Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(
                MyLinqDir + "Table with Text Colors Applied to Columns Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Text Colors Applied to Columns Report.docx");
            //ExEnd:ApplyingTextColorsToTableColumns

            // Test the report.
            CompareDocs("Table with Text Colors Applied to Columns Report.docx",
                "Table with Text Colors Applied to Columns Report Gold.docx");
        }

        [Test]
        public void MergingTableCellsVertically()
        {
            //ExStart:MergingTableCellsVertically
            //GistId:0d477de87c617d7dbbe4b710a0271568
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Cells Merged Vertically Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Table with Cells Merged Vertically Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Cells Merged Vertically Report.docx");
            //ExEnd:MergingTableCellsVertically

            // Test the report.
            CompareDocs("Table with Cells Merged Vertically Report.docx",
                "Table with Cells Merged Vertically Report Gold.docx");
        }

        [Test]
        public void MergingTableCellsHorizontally()
        {
            //ExStart:MergingTableCellsHorizontally
            //GistId:be43c7e51ac9138792a43c08d599e154
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Cells Merged Horizontally Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Table with Cells Merged Horizontally Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Cells Merged Horizontally Report.docx");
            //ExEnd:MergingTableCellsHorizontally

            // Test the report.
            CompareDocs("Table with Cells Merged Horizontally Report.docx",
                "Table with Cells Merged Horizontally Report Gold.docx");
        }

        [Test]
        public void MergingTableCellsVerticallyAndHorizontally()
        {
            //ExStart:MergingTableCellsVerticallyAndHorizontally
            //GistId:fd21a16d172df2129f52303978551c5c
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Cells Merged Vertically and Horizontally Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(
                MyLinqDir + "Table with Cells Merged Vertically and Horizontally Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Cells Merged Vertically and Horizontally Report.docx");
            //ExEnd:MergingTableCellsVerticallyAndHorizontally

            // Test the report.
            CompareDocs("Table with Cells Merged Vertically and Horizontally Report.docx",
                "Table with Cells Merged Vertically and Horizontally Report Gold.docx");
        }

        [Test]
        public void RestrictingTableCellMerging()
        {
            //ExStart:RestrictingTableCellMerging
            //GistId:ddf6b9088880b166c2619841b41bbc76
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Table with Cell Merging Restriction Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Table with Cell Merging Restriction Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "items");

            // Save the report.
            doc.Save(ArtifactsDir + "Table with Cell Merging Restriction Report.docx");
            //ExEnd:RestrictingTableCellMerging

            // Test the report.
            CompareDocs("Table with Cell Merging Restriction Report.docx",
                "Table with Cell Merging Restriction Report Gold.docx");
        }

        [Test]
        public void BuildingCrossTable()
        {
            //ExStart:BuildingCrossTable
            //GistId:7f70b373bd6cd2a7e6a3d2be56e8adce
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Cross Table Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Cross Table Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "ds");

            // Save the report.
            doc.Save(ArtifactsDir + "Cross Table Report.docx");
            //ExEnd:BuildingCrossTable

            // Test the report.
            CompareDocs("Cross Table Report.docx", "Cross Table Report Gold.docx");
        }

        [Test]
        public void BuildingCrossTableWithTotals()
        {
            //ExStart:BuildingCrossTableWithTotals
            //GistId:7b543cd536b61ee8205f9d546b3230e9
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Cross Table with Totals Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Cross Table with Totals Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "ds");

            // Save the report.
            doc.Save(ArtifactsDir + "Cross Table with Totals Report.docx");
            //ExEnd:BuildingCrossTableWithTotals

            // Test the report.
            CompareDocs("Cross Table with Totals Report.docx", "Cross Table with Totals Report Gold.docx");
        }

        [Test]
        public void BuildingCrossTableWithMergedCells()
        {
            //ExStart:BuildingCrossTableWithMergedCells
            //GistId:634e8f47cd509bca96c2a8a7e834ef7e
            // Open the template document.
            Document doc = new Document(MyLinqDir + "Cross Table with Merged Cells Template.docx");

            // Open the data source file.
            JsonDataSource dataSource = new JsonDataSource(MyLinqDir + "Cross Table with Merged Cells Data.json");

            // Build a report. The name of the data source should match the one used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.Options |= ReportBuildOptions.RemoveEmptyParagraphs; // Needed to remove extra empty paragraphs.
            engine.BuildReport(doc, dataSource, "ds");

            // Save the report.
            doc.Save(ArtifactsDir + "Cross Table with Merged Cells Report.docx");
            //ExEnd:BuildingCrossTableWithMergedCells

            // Test the report.
            CompareDocs("Cross Table with Merged Cells Report.docx", "Cross Table with Merged Cells Report Gold.docx");
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
