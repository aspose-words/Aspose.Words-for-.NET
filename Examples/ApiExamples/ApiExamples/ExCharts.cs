// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;

namespace ApiExamples
{
    [TestFixture]
    public class ExCharts : ApiExampleBase
    {
        [Test]
        public void ChartTitle()
        {
            //ExStart:ChartTitle
            //GistId:3428e84add5beb0d46a8face6e5fc858
            //ExFor:Chart
            //ExFor:Chart.Title
            //ExFor:ChartTitle
            //ExFor:ChartTitle.Overlay
            //ExFor:ChartTitle.Show
            //ExFor:ChartTitle.Text
            //ExFor:ChartTitle.Font
            //ExSummary:Shows how to insert a chart and set a title.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a chart shape with a document builder and get its chart.
            Shape chartShape = builder.InsertChart(ChartType.Bar, 400, 300);
            Chart chart = chartShape.Chart;

            // Use the "Title" property to give our chart a title, which appears at the top center of the chart area.
            ChartTitle title = chart.Title;
            title.Text = "My Chart";
            title.Font.Size = 15;
            title.Font.Color = Color.Blue;

            // Set the "Show" property to "true" to make the title visible. 
            title.Show = true;

            // Set the "Overlay" property to "true" Give other chart elements more room by allowing them to overlap the title
            title.Overlay = true;

            doc.Save(ArtifactsDir + "Charts.ChartTitle.docx");
            //ExEnd:ChartTitle

            doc = new Document(ArtifactsDir + "Charts.ChartTitle.docx");
            chartShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual(ShapeType.NonPrimitive, chartShape.ShapeType);
            Assert.True(chartShape.HasChart);

            title = chartShape.Chart.Title;

            Assert.AreEqual("My Chart", title.Text);
            Assert.True(title.Overlay);
            Assert.True(title.Show);
        }

        [Test]
        public void DataLabelNumberFormat()
        {
            //ExStart
            //ExFor:ChartDataLabelCollection.NumberFormat
            //ExFor:ChartDataLabelCollection.Font
            //ExFor:ChartNumberFormat.FormatCode
            //ExSummary:Shows how to enable and configure data labels for a chart series.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a line chart, then clear its demo data series to start with a clean chart,
            // and then set a title.
            Shape shape = builder.InsertChart(ChartType.Line, 500, 300);
            Chart chart = shape.Chart;
            chart.Series.Clear();
            chart.Title.Text = "Monthly sales report";

            // Insert a custom chart series with months as categories for the X-axis,
            // and respective decimal amounts for the Y-axis.
            ChartSeries series = chart.Series.Add("Revenue",
                new[] { "January", "February", "March" },
                new[] { 25.611d, 21.439d, 33.750d });

            // Enable data labels, and then apply a custom number format for values displayed in the data labels.
            // This format will treat displayed decimal values as millions of US Dollars.
            series.HasDataLabels = true;
            ChartDataLabelCollection dataLabels = series.DataLabels;
            dataLabels.ShowValue = true;
            dataLabels.NumberFormat.FormatCode = "\"US$\" #,##0.000\"M\"";
            dataLabels.Font.Size = 12;

            doc.Save(ArtifactsDir + "Charts.DataLabelNumberFormat.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.DataLabelNumberFormat.docx");
            series = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart.Series[0];

            Assert.True(series.HasDataLabels);
            Assert.True(series.DataLabels.ShowValue);
            Assert.AreEqual("\"US$\" #,##0.000\"M\"", series.DataLabels.NumberFormat.FormatCode);
        }

        [Test]
        public void AxisProperties()
        {
            //ExStart
            //ExFor:ChartAxis
            //ExFor:ChartAxis.CategoryType
            //ExFor:ChartAxis.Crosses
            //ExFor:ChartAxis.ReverseOrder
            //ExFor:ChartAxis.MajorTickMark
            //ExFor:ChartAxis.MinorTickMark
            //ExFor:ChartAxis.MajorUnit
            //ExFor:ChartAxis.MinorUnit
            //ExFor:AxisTickLabels
            //ExFor:AxisTickLabels.Offset
            //ExFor:AxisTickLabels.Position
            //ExFor:AxisTickLabels.IsAutoSpacing
            //ExFor:AxisTickLabels.Alignment
            //ExFor:AxisTickLabels.Font
            //ExFor:AxisTickLabels.Spacing
            //ExFor:ChartAxis.TickMarkSpacing
            //ExFor:AxisCategoryType
            //ExFor:AxisCrosses
            //ExFor:Chart.AxisX
            //ExFor:Chart.AxisY
            //ExFor:Chart.AxisZ
            //ExSummary:Shows how to insert a chart and modify the appearance of its axes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Column, 500, 300);
            Chart chart = shape.Chart;

            // Clear the chart's demo data series to start with a clean chart.
            chart.Series.Clear();

            // Insert a chart series with categories for the X-axis and respective numeric values for the Y-axis.
            chart.Series.Add("Aspose Test Series",
                new[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
                new double[] { 640, 320, 280, 120, 150 });

            // Chart axes have various options that can change their appearance,
            // such as their direction, major/minor unit ticks, and tick marks.
            ChartAxis xAxis = chart.AxisX;
            xAxis.CategoryType = AxisCategoryType.Category;
            xAxis.Crosses = AxisCrosses.Minimum;
            xAxis.ReverseOrder = false;
            xAxis.MajorTickMark = AxisTickMark.Inside;
            xAxis.MinorTickMark = AxisTickMark.Cross;
            xAxis.MajorUnit = 10.0d;
            xAxis.MinorUnit = 15.0d;
            xAxis.TickLabels.Offset = 50;
            xAxis.TickLabels.Position = AxisTickLabelPosition.Low;
            xAxis.TickLabels.IsAutoSpacing = false;
            xAxis.TickMarkSpacing = 1;

            ChartAxis yAxis = chart.AxisY;
            yAxis.CategoryType = AxisCategoryType.Automatic;
            yAxis.Crosses = AxisCrosses.Maximum;
            yAxis.ReverseOrder = true;
            yAxis.MajorTickMark = AxisTickMark.Inside;
            yAxis.MinorTickMark = AxisTickMark.Cross;
            yAxis.MajorUnit = 100.0d;
            yAxis.MinorUnit = 20.0d;
            yAxis.TickLabels.Position = AxisTickLabelPosition.NextToAxis;
            yAxis.TickLabels.Alignment = ParagraphAlignment.Center;
            yAxis.TickLabels.Font.Color = Color.Red;
            yAxis.TickLabels.Spacing = 1;

            // Column charts do not have a Z-axis.
            Assert.Null(chart.AxisZ);

            doc.Save(ArtifactsDir + "Charts.AxisProperties.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.AxisProperties.docx");
            chart = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart;

            Assert.AreEqual(AxisCategoryType.Category, chart.AxisX.CategoryType);
            Assert.AreEqual(AxisCrosses.Minimum, chart.AxisX.Crosses);
            Assert.False(chart.AxisX.ReverseOrder);
            Assert.AreEqual(AxisTickMark.Inside, chart.AxisX.MajorTickMark);
            Assert.AreEqual(AxisTickMark.Cross, chart.AxisX.MinorTickMark);
            Assert.AreEqual(1.0d, chart.AxisX.MajorUnit);
            Assert.AreEqual(0.5d, chart.AxisX.MinorUnit);
            Assert.AreEqual(50, chart.AxisX.TickLabels.Offset);
            Assert.AreEqual(AxisTickLabelPosition.Low, chart.AxisX.TickLabels.Position);
            Assert.False(chart.AxisX.TickLabels.IsAutoSpacing);
            Assert.AreEqual(1, chart.AxisX.TickMarkSpacing);

            Assert.AreEqual(AxisCategoryType.Category, chart.AxisY.CategoryType);
            Assert.AreEqual(AxisCrosses.Maximum, chart.AxisY.Crosses);
            Assert.True(chart.AxisY.ReverseOrder);
            Assert.AreEqual(AxisTickMark.Inside, chart.AxisY.MajorTickMark);
            Assert.AreEqual(AxisTickMark.Cross, chart.AxisY.MinorTickMark);
            Assert.AreEqual(100.0d, chart.AxisY.MajorUnit);
            Assert.AreEqual(20.0d, chart.AxisY.MinorUnit);
            Assert.AreEqual(AxisTickLabelPosition.NextToAxis, chart.AxisY.TickLabels.Position);
            Assert.AreEqual(ParagraphAlignment.Center, chart.AxisY.TickLabels.Alignment);
            Assert.AreEqual(Color.Red.ToArgb(), chart.AxisY.TickLabels.Font.Color.ToArgb());
            Assert.AreEqual(1, chart.AxisY.TickLabels.Spacing);
        }

        [Test]
        public void AxisCollection()
        {
            //ExStart
            //ExFor:ChartAxisCollection
            //ExFor:Chart.Axes
            //ExSummary:Shows how to work with axes collection.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Column, 500, 300);
            Chart chart = shape.Chart;

            // Hide the major grid lines on the primary and secondary Y axes.
            foreach (ChartAxis axis in chart.Axes)
            {
                if (axis.Type == ChartAxisType.Value)
                    axis.HasMajorGridlines = false;
            }

            doc.Save(ArtifactsDir + "Charts.AxisCollection.docx");
            //ExEnd
        }

        [Test]
        public void DateTimeValues()
        {
            //ExStart
            //ExFor:AxisBound
            //ExFor:AxisBound.#ctor(Double)
            //ExFor:AxisBound.#ctor(DateTime)
            //ExFor:AxisScaling.Minimum
            //ExFor:AxisScaling.Maximum
            //ExFor:ChartAxis.Scaling
            //ExFor:AxisTickMark
            //ExFor:AxisTickLabelPosition
            //ExFor:AxisTimeUnit
            //ExFor:ChartAxis.BaseTimeUnit
            //ExFor:ChartAxis.HasMajorGridlines
            //ExFor:ChartAxis.HasMinorGridlines
            //ExSummary:Shows how to insert chart with date/time values.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Line, 500, 300);
            Chart chart = shape.Chart;

            // Clear the chart's demo data series to start with a clean chart.
            chart.Series.Clear();

            // Add a custom series containing date/time values for the X-axis, and respective decimal values for the Y-axis.
            chart.Series.Add("Aspose Test Series",
                new[]
                {
                    new DateTime(2017, 11, 06), new DateTime(2017, 11, 09), new DateTime(2017, 11, 15),
                    new DateTime(2017, 11, 21), new DateTime(2017, 11, 25), new DateTime(2017, 11, 29)
                },
                new[] { 1.2, 0.3, 2.1, 2.9, 4.2, 5.3 });


            // Set lower and upper bounds for the X-axis.
            ChartAxis xAxis = chart.AxisX;
            xAxis.Scaling.Minimum = new AxisBound(new DateTime(2017, 11, 05).ToOADate());
            xAxis.Scaling.Maximum = new AxisBound(new DateTime(2017, 12, 03));

            // Set the major units of the X-axis to a week, and the minor units to a day.
            xAxis.BaseTimeUnit = AxisTimeUnit.Days;
            xAxis.MajorUnit = 7.0d;
            xAxis.MajorTickMark = AxisTickMark.Cross;
            xAxis.MinorUnit = 1.0d;
            xAxis.MinorTickMark = AxisTickMark.Outside;
            xAxis.HasMajorGridlines = true;
            xAxis.HasMinorGridlines = true;

            // Define Y-axis properties for decimal values.
            ChartAxis yAxis = chart.AxisY;
            yAxis.TickLabels.Position = AxisTickLabelPosition.High;
            yAxis.MajorUnit = 100.0d;
            yAxis.MinorUnit = 50.0d;
            yAxis.DisplayUnit.Unit = AxisBuiltInUnit.Hundreds;
            yAxis.Scaling.Minimum = new AxisBound(100);
            yAxis.Scaling.Maximum = new AxisBound(700);
            yAxis.HasMajorGridlines = true;
            yAxis.HasMinorGridlines = true;

            doc.Save(ArtifactsDir + "Charts.DateTimeValues.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.DateTimeValues.docx");
            chart = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart;

            Assert.AreEqual(new AxisBound(new DateTime(2017, 11, 05).ToOADate()), chart.AxisX.Scaling.Minimum);
            Assert.AreEqual(new AxisBound(new DateTime(2017, 12, 03)), chart.AxisX.Scaling.Maximum);
            Assert.AreEqual(AxisTimeUnit.Days, chart.AxisX.BaseTimeUnit);
            Assert.AreEqual(7.0d, chart.AxisX.MajorUnit);
            Assert.AreEqual(1.0d, chart.AxisX.MinorUnit);
            Assert.AreEqual(AxisTickMark.Cross, chart.AxisX.MajorTickMark);
            Assert.AreEqual(AxisTickMark.Outside, chart.AxisX.MinorTickMark);
            Assert.AreEqual(true, chart.AxisX.HasMajorGridlines);
            Assert.AreEqual(true, chart.AxisX.HasMinorGridlines);

            Assert.AreEqual(AxisTickLabelPosition.High, chart.AxisY.TickLabels.Position);
            Assert.AreEqual(100.0d, chart.AxisY.MajorUnit);
            Assert.AreEqual(50.0d, chart.AxisY.MinorUnit);
            Assert.AreEqual(AxisBuiltInUnit.Hundreds, chart.AxisY.DisplayUnit.Unit);
            Assert.AreEqual(new AxisBound(100), chart.AxisY.Scaling.Minimum);
            Assert.AreEqual(new AxisBound(700), chart.AxisY.Scaling.Maximum);
            Assert.AreEqual(true, chart.AxisY.HasMajorGridlines);
            Assert.AreEqual(true, chart.AxisY.HasMinorGridlines);
        }

        [Test]
        public void HideChartAxis()
        {
            //ExStart
            //ExFor:ChartAxis.Hidden
            //ExSummary:Shows how to hide chart axes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Line, 500, 300);
            Chart chart = shape.Chart;

            // Clear the chart's demo data series to start with a clean chart.
            chart.Series.Clear();

            // Add a custom series with categories for the X-axis, and respective decimal values for the Y-axis.
            chart.Series.Add("AW Series 1",
                new[] { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" },
                new[] { 1.2, 0.3, 2.1, 2.9, 4.2 });

            // Hide the chart axes to simplify the appearance of the chart. 
            chart.AxisX.Hidden = true;
            chart.AxisY.Hidden = true;

            doc.Save(ArtifactsDir + "Charts.HideChartAxis.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.HideChartAxis.docx");
            chart = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart;

            Assert.True(chart.AxisX.Hidden);
            Assert.True(chart.AxisY.Hidden);
        }

        [Test]
        public void SetNumberFormatToChartAxis()
        {
            //ExStart
            //ExFor:ChartAxis.NumberFormat
            //ExFor:ChartNumberFormat
            //ExFor:ChartNumberFormat.FormatCode
            //ExFor:ChartNumberFormat.IsLinkedToSource
            //ExSummary:Shows how to set formatting for chart values.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Column, 500, 300);
            Chart chart = shape.Chart;

            // Clear the chart's demo data series to start with a clean chart.
            chart.Series.Clear();

            // Add a custom series to the chart with categories for the X-axis,
            // and large respective numeric values for the Y-axis. 
            chart.Series.Add("Aspose Test Series",
                new[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
                new double[] { 1900000, 850000, 2100000, 600000, 1500000 });

            // Set the number format of the Y-axis tick labels to not group digits with commas. 
            chart.AxisY.NumberFormat.FormatCode = "#,##0";

            // This flag can override the above value and draw the number format from the source cell.
            Assert.False(chart.AxisY.NumberFormat.IsLinkedToSource);

            doc.Save(ArtifactsDir + "Charts.SetNumberFormatToChartAxis.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.SetNumberFormatToChartAxis.docx");
            chart = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart;

            Assert.AreEqual("#,##0", chart.AxisY.NumberFormat.FormatCode);
        }

        [TestCase(ChartType.Column)]
        [TestCase(ChartType.Line)]
        [TestCase(ChartType.Pie)]
        [TestCase(ChartType.Bar)]
        [TestCase(ChartType.Area)]
        public void TestDisplayChartsWithConversion(ChartType chartType)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(chartType, 500, 300);
            Chart chart = shape.Chart;
            chart.Series.Clear();
            
            chart.Series.Add("Aspose Test Series",
                new[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
                new double[] { 1900000, 850000, 2100000, 600000, 1500000 });

            doc.Save(ArtifactsDir + "Charts.TestDisplayChartsWithConversion.docx");
            doc.Save(ArtifactsDir + "Charts.TestDisplayChartsWithConversion.pdf");
        }

        [Test]
        public void Surface3DChart()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Surface3D, 500, 300);
            Chart chart = shape.Chart;
            chart.Series.Clear();

            chart.Series.Add("Aspose Test Series 1",
                new string[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
                new double[] { 1900000, 850000, 2100000, 600000, 1500000 });

            chart.Series.Add("Aspose Test Series 2",
                new string[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
                new double[] { 900000, 50000, 1100000, 400000, 2500000 });

            chart.Series.Add("Aspose Test Series 3",
                new string[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
                new double[] { 500000, 820000, 1500000, 400000, 100000 });

            doc.Save(ArtifactsDir + "Charts.SurfaceChart.docx");
            doc.Save(ArtifactsDir + "Charts.SurfaceChart.pdf");
        }

        [Test]
        public void DataLabelsBubbleChart()
        {
            //ExStart
            //ExFor:ChartDataLabelCollection.Separator
            //ExFor:ChartDataLabelCollection.ShowBubbleSize
            //ExFor:ChartDataLabelCollection.ShowCategoryName
            //ExFor:ChartDataLabelCollection.ShowSeriesName
            //ExSummary:Shows how to work with data labels of a bubble chart.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Chart chart = builder.InsertChart(ChartType.Bubble, 500, 300).Chart;

            // Clear the chart's demo data series to start with a clean chart.
            chart.Series.Clear();

            // Add a custom series with X/Y coordinates and diameter of each of the bubbles. 
            ChartSeries series = chart.Series.Add("Aspose Test Series",
                new[] { 2.9, 3.5, 1.1, 4.0, 4.0 },
                new[] { 1.9, 8.5, 2.1, 6.0, 1.5 },
                new[] { 9.0, 4.5, 2.5, 8.0, 5.0 });

            // Enable data labels, and then modify their appearance.
            series.HasDataLabels = true;
            ChartDataLabelCollection dataLabels = series.DataLabels;
            dataLabels.ShowBubbleSize = true;
            dataLabels.ShowCategoryName = true;
            dataLabels.ShowSeriesName = true;
            dataLabels.Separator = " & ";

            doc.Save(ArtifactsDir + "Charts.DataLabelsBubbleChart.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.DataLabelsBubbleChart.docx");
            dataLabels = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart.Series[0].DataLabels;

            Assert.True(dataLabels.ShowBubbleSize);
            Assert.True(dataLabels.ShowCategoryName);
            Assert.True(dataLabels.ShowSeriesName);
            Assert.AreEqual(" & ", dataLabels.Separator);
        }

        [Test]
        public void DataLabelsPieChart()
        {
            //ExStart
            //ExFor:ChartDataLabelCollection.Separator
            //ExFor:ChartDataLabelCollection.ShowLeaderLines
            //ExFor:ChartDataLabelCollection.ShowLegendKey
            //ExFor:ChartDataLabelCollection.ShowPercentage
            //ExFor:ChartDataLabelCollection.ShowValue
            //ExSummary:Shows how to work with data labels of a pie chart.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Chart chart = builder.InsertChart(ChartType.Pie, 500, 300).Chart;

            // Clear the chart's demo data series to start with a clean chart.
            chart.Series.Clear();

            // Insert a custom chart series with a category name for each of the sectors, and their frequency table.
            ChartSeries series = chart.Series.Add("Aspose Test Series",
                new[] { "Word", "PDF", "Excel" },
                new[] { 2.7, 3.2, 0.8 });

            // Enable data labels that will display both percentage and frequency of each sector, and modify their appearance.
            series.HasDataLabels = true;
            ChartDataLabelCollection dataLabels = series.DataLabels;
            dataLabels.ShowLeaderLines = true;
            dataLabels.ShowLegendKey = true;
            dataLabels.ShowPercentage = true;
            dataLabels.ShowValue = true;
            dataLabels.Separator = "; ";

            doc.Save(ArtifactsDir + "Charts.DataLabelsPieChart.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.DataLabelsPieChart.docx");
            dataLabels = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart.Series[0].DataLabels;

            Assert.True(dataLabels.ShowLeaderLines);
            Assert.True(dataLabels.ShowLegendKey);
            Assert.True(dataLabels.ShowPercentage);
            Assert.True(dataLabels.ShowValue);
            Assert.AreEqual("; ", dataLabels.Separator);
        }

        //ExStart
        //ExFor:ChartSeries
        //ExFor:ChartSeries.DataLabels
        //ExFor:ChartSeries.DataPoints
        //ExFor:ChartSeries.Name
        //ExFor:ChartDataLabel
        //ExFor:ChartDataLabel.Index
        //ExFor:ChartDataLabel.IsVisible
        //ExFor:ChartDataLabel.NumberFormat
        //ExFor:ChartDataLabel.Separator
        //ExFor:ChartDataLabel.ShowCategoryName
        //ExFor:ChartDataLabel.ShowDataLabelsRange
        //ExFor:ChartDataLabel.ShowLeaderLines
        //ExFor:ChartDataLabel.ShowLegendKey
        //ExFor:ChartDataLabel.ShowPercentage
        //ExFor:ChartDataLabel.ShowSeriesName
        //ExFor:ChartDataLabel.ShowValue
        //ExFor:ChartDataLabel.IsHidden
        //ExFor:ChartDataLabelCollection
        //ExFor:ChartDataLabelCollection.ClearFormat
        //ExFor:ChartDataLabelCollection.Count
        //ExFor:ChartDataLabelCollection.GetEnumerator
        //ExFor:ChartDataLabelCollection.Item(Int32)
        //ExSummary:Shows how to apply labels to data points in a line chart.
        [Test] //ExSkip
        public void DataLabels()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape chartShape = builder.InsertChart(ChartType.Line, 400, 300);
            Chart chart = chartShape.Chart;

            Assert.AreEqual(3, chart.Series.Count);
            Assert.AreEqual("Series 1", chart.Series[0].Name);
            Assert.AreEqual("Series 2", chart.Series[1].Name);
            Assert.AreEqual("Series 3", chart.Series[2].Name);

            // Apply data labels to every series in the chart.
            // These labels will appear next to each data point in the graph and display its value.
            foreach (ChartSeries series in chart.Series)
            {
                ApplyDataLabels(series, 4, "000.0", ", ");
                Assert.AreEqual(4, series.DataLabels.Count);
            }

            // Change the separator string for every data label in a series.
            using (IEnumerator<ChartDataLabel> enumerator = chart.Series[0].DataLabels.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    Assert.AreEqual(", ", enumerator.Current.Separator);
                    enumerator.Current.Separator = " & ";
                }
            }

            // For a cleaner looking graph, we can remove data labels individually.
            chart.Series[1].DataLabels[2].ClearFormat();

            // We can also strip an entire series of its data labels at once.
            chart.Series[2].DataLabels.ClearFormat();

            doc.Save(ArtifactsDir + "Charts.DataLabels.docx");
        }

        /// <summary>
        /// Apply data labels with custom number format and separator to several data points in a series.
        /// </summary>
        private static void ApplyDataLabels(ChartSeries series, int labelsCount, string numberFormat, string separator)
        {
            for (int i = 0; i < labelsCount; i++)
            {
                series.HasDataLabels = true;

                Assert.False(series.DataLabels[i].IsVisible);

                series.DataLabels[i].ShowCategoryName = true;
                series.DataLabels[i].ShowSeriesName = true;
                series.DataLabels[i].ShowValue = true;
                series.DataLabels[i].ShowLeaderLines = true;
                series.DataLabels[i].ShowLegendKey = true;
                series.DataLabels[i].ShowPercentage = false;
                series.DataLabels[i].IsHidden = false;
                Assert.False(series.DataLabels[i].ShowDataLabelsRange);

                series.DataLabels[i].NumberFormat.FormatCode = numberFormat;
                series.DataLabels[i].Separator = separator;

                Assert.False(series.DataLabels[i].ShowDataLabelsRange);
                Assert.True(series.DataLabels[i].IsVisible);
                Assert.False(series.DataLabels[i].IsHidden);
            }
        }
        //ExEnd

        //ExStart
        //ExFor:ChartSeries.Smooth
        //ExFor:ChartDataPoint
        //ExFor:ChartDataPoint.Index
        //ExFor:ChartDataPointCollection
        //ExFor:ChartDataPointCollection.ClearFormat
        //ExFor:ChartDataPointCollection.Count
        //ExFor:ChartDataPointCollection.GetEnumerator
        //ExFor:ChartDataPointCollection.Item(Int32)
        //ExFor:ChartMarker
        //ExFor:ChartMarker.Size
        //ExFor:ChartMarker.Symbol
        //ExFor:IChartDataPoint
        //ExFor:IChartDataPoint.InvertIfNegative
        //ExFor:IChartDataPoint.Marker
        //ExFor:MarkerSymbol
        //ExSummary:Shows how to work with data points on a line chart.
        [Test]//ExSkip
        public void ChartDataPoint()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Line, 500, 350);
            Chart chart = shape.Chart;

            Assert.AreEqual(3, chart.Series.Count);
            Assert.AreEqual("Series 1", chart.Series[0].Name);
            Assert.AreEqual("Series 2", chart.Series[1].Name);
            Assert.AreEqual("Series 3", chart.Series[2].Name);

            // Emphasize the chart's data points by making them appear as diamond shapes.
            foreach (ChartSeries series in chart.Series)
                ApplyDataPoints(series, 4, MarkerSymbol.Diamond, 15);

            // Smooth out the line that represents the first data series.
            chart.Series[0].Smooth = true;

            // Verify that data points for the first series will not invert their colors if the value is negative.
            using (IEnumerator<ChartDataPoint> enumerator = chart.Series[0].DataPoints.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    Assert.False(enumerator.Current.InvertIfNegative);
                }
            }

            // For a cleaner looking graph, we can clear format individually.
            chart.Series[1].DataPoints[2].ClearFormat();

            // We can also strip an entire series of data points at once.
            chart.Series[2].DataPoints.ClearFormat();

            doc.Save(ArtifactsDir + "Charts.ChartDataPoint.docx");
        }

        /// <summary>
        /// Applies a number of data points to a series.
        /// </summary>
        private static void ApplyDataPoints(ChartSeries series, int dataPointsCount, MarkerSymbol markerSymbol, int dataPointSize)
        {
            for (int i = 0; i < dataPointsCount; i++)
            {
                ChartDataPoint point = series.DataPoints[i];
                point.Marker.Symbol = markerSymbol;
                point.Marker.Size = dataPointSize;

                Assert.AreEqual(i, point.Index);
            }
        }
        //ExEnd

        [Test]
        public void PieChartExplosion()
        {
            //ExStart
            //ExFor:IChartDataPoint.Explosion
            //ExSummary:Shows how to move the slices of a pie chart away from the center.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Pie, 500, 350);
            Chart chart = shape.Chart;

            Assert.AreEqual(1, chart.Series.Count);
            Assert.AreEqual("Sales", chart.Series[0].Name);

            // "Slices" of a pie chart may be moved away from the center by a distance via the respective data point's Explosion attribute.
            // Add a data point to the first portion of the pie chart and move it away from the center by 10 points.
            // Aspose.Words create data points automatically if them does not exist.
            ChartDataPoint dataPoint = chart.Series[0].DataPoints[0];
            dataPoint.Explosion = 10;

            // Displace the second portion by a greater distance.
            dataPoint = chart.Series[0].DataPoints[1];
            dataPoint.Explosion = 40;

            doc.Save(ArtifactsDir + "Charts.PieChartExplosion.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.PieChartExplosion.docx");
            ChartSeries series = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart.Series[0];

            Assert.AreEqual(10, series.DataPoints[0].Explosion);
            Assert.AreEqual(40, series.DataPoints[1].Explosion);
        }

        [Test]
        public void Bubble3D()
        {
            //ExStart
            //ExFor:ChartDataLabel.ShowBubbleSize
            //ExFor:ChartDataLabel.Font
            //ExFor:IChartDataPoint.Bubble3D
            //ExSummary:Shows how to use 3D effects with bubble charts.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Bubble3D, 500, 350);
            Chart chart = shape.Chart;

            Assert.AreEqual(1, chart.Series.Count);
            Assert.AreEqual("Y-Values", chart.Series[0].Name);
            Assert.True(chart.Series[0].Bubble3D);

            // Apply a data label to each bubble that displays its diameter.
            for (int i = 0; i < 3; i++)
            {
                chart.Series[0].HasDataLabels = true;
                chart.Series[0].DataLabels[i].ShowBubbleSize = true;
                chart.Series[0].DataLabels[i].Font.Size = 12;
            }

            doc.Save(ArtifactsDir + "Charts.Bubble3D.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.Bubble3D.docx");
            ChartSeries series = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart.Series[0];

            for (int i = 0; i < 3; i++)
            {
                Assert.True(series.DataLabels[i].ShowBubbleSize);
            }
        }

        //ExStart
        //ExFor:ChartAxis.Type
        //ExFor:ChartAxisType
        //ExFor:ChartType
        //ExFor:Chart.Series
        //ExFor:ChartSeriesCollection.Add(String,DateTime[],Double[])
        //ExFor:ChartSeriesCollection.Add(String,Double[],Double[])
        //ExFor:ChartSeriesCollection.Add(String,Double[],Double[],Double[])
        //ExFor:ChartSeriesCollection.Add(String,String[],Double[])
        //ExSummary:Shows how to create an appropriate type of chart series for a graph type.
        [Test] //ExSkip
        public void ChartSeriesCollection()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // There are several ways of populating a chart's series collection.
            // Different series schemas are intended for different chart types.
            // 1 -  Column chart with columns grouped and banded along the X-axis by category:
            Chart chart = AppendChart(builder, ChartType.Column, 500, 300);

            string[] categories = { "Category 1", "Category 2", "Category 3" };

            // Insert two series of decimal values containing a value for each respective category.
            // This column chart will have three groups, each with two columns.
            chart.Series.Add("Series 1", categories, new[] { 76.6, 82.1, 91.6 });
            chart.Series.Add("Series 2", categories, new[] { 64.2, 79.5, 94.0 });

            // Categories are distributed along the X-axis, and values are distributed along the Y-axis.
            Assert.AreEqual(ChartAxisType.Category, chart.AxisX.Type);
            Assert.AreEqual(ChartAxisType.Value, chart.AxisY.Type);

            // 2 -  Area chart with dates distributed along the X-axis:
            chart = AppendChart(builder, ChartType.Area, 500, 300);

            DateTime[] dates = { new DateTime(2014, 3, 31),
                new DateTime(2017, 1, 23),
                new DateTime(2017, 6, 18),
                new DateTime(2019, 11, 22),
                new DateTime(2020, 9, 7)
            };

            // Insert a series with a decimal value for each respective date.
            // The dates will be distributed along a linear X-axis,
            // and the values added to this series will create data points.
            chart.Series.Add("Series 1", dates, new[] { 15.8, 21.5, 22.9, 28.7, 33.1 });

            Assert.AreEqual(ChartAxisType.Category, chart.AxisX.Type);
            Assert.AreEqual(ChartAxisType.Value, chart.AxisY.Type);

            // 3 -  2D scatter plot:
            chart = AppendChart(builder, ChartType.Scatter, 500, 300);

            // Each series will need two decimal arrays of equal length.
            // The first array contains X-values, and the second contains corresponding Y-values
            // of data points on the chart's graph.
            chart.Series.Add("Series 1",
                new[] { 3.1, 3.5, 6.3, 4.1, 2.2, 8.3, 1.2, 3.6 },
                new[] { 3.1, 6.3, 4.6, 0.9, 8.5, 4.2, 2.3, 9.9 });
            chart.Series.Add("Series 2",
                new[] { 2.6, 7.3, 4.5, 6.6, 2.1, 9.3, 0.7, 3.3 },
                new[] { 7.1, 6.6, 3.5, 7.8, 7.7, 9.5, 1.3, 4.6 });

            Assert.AreEqual(ChartAxisType.Value, chart.AxisX.Type);
            Assert.AreEqual(ChartAxisType.Value, chart.AxisY.Type);

            // 4 -  Bubble chart:
            chart = AppendChart(builder, ChartType.Bubble, 500, 300);

            // Each series will need three decimal arrays of equal length.
            // The first array contains X-values, the second contains corresponding Y-values,
            // and the third contains diameters for each of the graph's data points.
            chart.Series.Add("Series 1",
                new[] { 1.1, 5.0, 9.8 },
                new[] { 1.2, 4.9, 9.9 },
                new[] { 2.0, 4.0, 8.0 });

            doc.Save(ArtifactsDir + "Charts.ChartSeriesCollection.docx");
        }

        /// <summary>
        /// Insert a chart using a document builder of a specified ChartType, width and height, and remove its demo data.
        /// </summary>
        private static Chart AppendChart(DocumentBuilder builder, ChartType chartType, double width, double height)
        {
            Shape chartShape = builder.InsertChart(chartType, width, height);
            Chart chart = chartShape.Chart;
            chart.Series.Clear();
            Assert.AreEqual(0, chart.Series.Count); //ExSkip

            return chart;
        }
        //ExEnd

        [Test]
        public void ChartSeriesCollectionModify()
        {
            //ExStart
            //ExFor:ChartSeriesCollection
            //ExFor:ChartSeriesCollection.Clear
            //ExFor:ChartSeriesCollection.Count
            //ExFor:ChartSeriesCollection.GetEnumerator
            //ExFor:ChartSeriesCollection.Item(Int32)
            //ExFor:ChartSeriesCollection.RemoveAt(Int32)
            //ExSummary:Shows how to add and remove series data in a chart.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a column chart that will contain three series of demo data by default.
            Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
            Chart chart = chartShape.Chart;

            // Each series has four decimal values: one for each of the four categories.
            // Four clusters of three columns will represent this data.
            ChartSeriesCollection chartData = chart.Series;

            Assert.AreEqual(3, chartData.Count);

            // Print the name of every series in the chart.
            using (IEnumerator<ChartSeries> enumerator = chart.Series.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    Console.WriteLine(enumerator.Current.Name);
                }
            }

            // These are the names of the categories in the chart.
            string[] categories = { "Category 1", "Category 2", "Category 3", "Category 4" };

            // We can add a series with new values for existing categories.
            // This chart will now contain four clusters of four columns.
            chart.Series.Add("Series 4", categories, new[] { 4.4, 7.0, 3.5, 2.1 });
            Assert.AreEqual(4, chartData.Count); //ExSkip
            Assert.AreEqual("Series 4", chartData[3].Name); //ExSkip

            // A chart series can also be removed by index, like this.
            // This will remove one of the three demo series that came with the chart.
            chartData.RemoveAt(2);

            Assert.False(chartData.Any(s => s.Name == "Series 3"));
            Assert.AreEqual(3, chartData.Count); //ExSkip
            Assert.AreEqual("Series 4", chartData[2].Name); //ExSkip

            // We can also clear all the chart's data at once with this method.
            // When creating a new chart, this is the way to wipe all the demo data
            // before we can begin working on a blank chart.
            chartData.Clear();
            Assert.AreEqual(0, chartData.Count); //ExSkip

            //ExEnd
        }

        [Test]
        public void AxisScaling()
        {
            //ExStart
            //ExFor:AxisScaleType
            //ExFor:AxisScaling
            //ExFor:AxisScaling.LogBase
            //ExFor:AxisScaling.Type
            //ExSummary:Shows how to apply logarithmic scaling to a chart axis.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape chartShape = builder.InsertChart(ChartType.Scatter, 450, 300);
            Chart chart = chartShape.Chart;

            // Clear the chart's demo data series to start with a clean chart.
            chart.Series.Clear();

            // Insert a series with X/Y coordinates for five points.
            chart.Series.Add("Series 1",
                new[] { 1.0, 2.0, 3.0, 4.0, 5.0 },
                new[] { 1.0, 20.0, 400.0, 8000.0, 160000.0 });

            // The scaling of the X-axis is linear by default,
            // displaying evenly incrementing values that cover our X-value range (0, 1, 2, 3...).
            // A linear axis is not ideal for our Y-values
            // since the points with the smaller Y-values will be harder to read.
            // A logarithmic scaling with a base of 20 (1, 20, 400, 8000...)
            // will spread the plotted points, allowing us to read their values on the chart more easily.
            chart.AxisY.Scaling.Type = AxisScaleType.Logarithmic;
            chart.AxisY.Scaling.LogBase = 20;

            doc.Save(ArtifactsDir + "Charts.AxisScaling.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.AxisScaling.docx");
            chart = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart;

            Assert.AreEqual(AxisScaleType.Linear, chart.AxisX.Scaling.Type);
            Assert.AreEqual(AxisScaleType.Logarithmic, chart.AxisY.Scaling.Type);
            Assert.AreEqual(20.0d, chart.AxisY.Scaling.LogBase);
        }

        [Test]
        public void AxisBound()
        {
            //ExStart
            //ExFor:AxisBound.#ctor
            //ExFor:AxisBound.IsAuto
            //ExFor:AxisBound.Value
            //ExFor:AxisBound.ValueAsDate
            //ExSummary:Shows how to set custom axis bounds.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape chartShape = builder.InsertChart(ChartType.Scatter, 450, 300);
            Chart chart = chartShape.Chart;

            // Clear the chart's demo data series to start with a clean chart.
            chart.Series.Clear();

            // Add a series with two decimal arrays. The first array contains the X-values,
            // and the second contains corresponding Y-values for points in the scatter chart.
            chart.Series.Add("Series 1",
                new[] { 1.1, 5.4, 7.9, 3.5, 2.1, 9.7 },
                new[] { 2.1, 0.3, 0.6, 3.3, 1.4, 1.9 });

            // By default, default scaling is applied to the graph's X and Y-axes,
            // so that both their ranges are big enough to encompass every X and Y-value of every series.
            Assert.True(chart.AxisX.Scaling.Minimum.IsAuto);

            // We can define our own axis bounds.
            // In this case, we will make both the X and Y-axis rulers show a range of 0 to 10.
            chart.AxisX.Scaling.Minimum = new AxisBound(0);
            chart.AxisX.Scaling.Maximum = new AxisBound(10);
            chart.AxisY.Scaling.Minimum = new AxisBound(0);
            chart.AxisY.Scaling.Maximum = new AxisBound(10);

            Assert.False(chart.AxisX.Scaling.Minimum.IsAuto);
            Assert.False(chart.AxisY.Scaling.Minimum.IsAuto);

            // Create a line chart with a series requiring a range of dates on the X-axis, and decimal values for the Y-axis.
            chartShape = builder.InsertChart(ChartType.Line, 450, 300);
            chart = chartShape.Chart;
            chart.Series.Clear();

            DateTime[] dates = { new DateTime(1973, 5, 11),
                new DateTime(1981, 2, 4),
                new DateTime(1985, 9, 23),
                new DateTime(1989, 6, 28),
                new DateTime(1994, 12, 15)
            };

            chart.Series.Add("Series 1", dates, new[] { 3.0, 4.7, 5.9, 7.1, 8.9 });

            // We can set axis bounds in the form of dates as well, limiting the chart to a period.
            // Setting the range to 1980-1990 will omit the two of the series values
            // that are outside of the range from the graph.
            chart.AxisX.Scaling.Minimum = new AxisBound(new DateTime(1980, 1, 1));
            chart.AxisX.Scaling.Maximum = new AxisBound(new DateTime(1990, 1, 1));

            doc.Save(ArtifactsDir + "Charts.AxisBound.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.AxisBound.docx");
            chart = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart;

            Assert.False(chart.AxisX.Scaling.Minimum.IsAuto);
            Assert.AreEqual(0.0d, chart.AxisX.Scaling.Minimum.Value);
            Assert.AreEqual(10.0d, chart.AxisX.Scaling.Maximum.Value);

            Assert.False(chart.AxisY.Scaling.Minimum.IsAuto);
            Assert.AreEqual(0.0d, chart.AxisY.Scaling.Minimum.Value);
            Assert.AreEqual(10.0d, chart.AxisY.Scaling.Maximum.Value);

            chart = ((Shape)doc.GetChild(NodeType.Shape, 1, true)).Chart;

            Assert.False(chart.AxisX.Scaling.Minimum.IsAuto);
            Assert.AreEqual(new AxisBound(new DateTime(1980, 1, 1)), chart.AxisX.Scaling.Minimum);
            Assert.AreEqual(new AxisBound(new DateTime(1990, 1, 1)), chart.AxisX.Scaling.Maximum);

            Assert.True(chart.AxisY.Scaling.Minimum.IsAuto);
        }

        [Test]
        public void ChartLegend()
        {
            //ExStart
            //ExFor:Chart.Legend
            //ExFor:ChartLegend
            //ExFor:ChartLegend.Overlay
            //ExFor:ChartLegend.Position
            //ExFor:LegendPosition
            //ExSummary:Shows how to edit the appearance of a chart's legend.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Line, 450, 300);
            Chart chart = shape.Chart;

            Assert.AreEqual(3, chart.Series.Count);
            Assert.AreEqual("Series 1", chart.Series[0].Name);
            Assert.AreEqual("Series 2", chart.Series[1].Name);
            Assert.AreEqual("Series 3", chart.Series[2].Name);

            // Move the chart's legend to the top right corner.
            ChartLegend legend = chart.Legend;
            legend.Position = LegendPosition.TopRight;

            // Give other chart elements, such as the graph, more room by allowing them to overlap the legend.
            legend.Overlay = true;

            doc.Save(ArtifactsDir + "Charts.ChartLegend.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.ChartLegend.docx");

            legend = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart.Legend;

            Assert.True(legend.Overlay);
            Assert.AreEqual(LegendPosition.TopRight, legend.Position);
        }

        [Test]
        public void AxisCross()
        {
            //ExStart
            //ExFor:ChartAxis.AxisBetweenCategories
            //ExFor:ChartAxis.CrossesAt
            //ExSummary:Shows how to get a graph axis to cross at a custom location.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Column, 450, 250);
            Chart chart = shape.Chart;

            Assert.AreEqual(3, chart.Series.Count);
            Assert.AreEqual("Series 1", chart.Series[0].Name);
            Assert.AreEqual("Series 2", chart.Series[1].Name);
            Assert.AreEqual("Series 3", chart.Series[2].Name);

            // For column charts, the Y-axis crosses at zero by default,
            // which means that columns for all values below zero point down to represent negative values.
            // We can set a different value for the Y-axis crossing. In this case, we will set it to 3.
            ChartAxis axis = chart.AxisX;
            axis.Crosses = AxisCrosses.Custom;
            axis.CrossesAt = 3;
            axis.AxisBetweenCategories = true;

            doc.Save(ArtifactsDir + "Charts.AxisCross.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.AxisCross.docx");
            axis = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart.AxisX;

            Assert.True(axis.AxisBetweenCategories);
            Assert.AreEqual(AxisCrosses.Custom, axis.Crosses);
            Assert.AreEqual(3.0d, axis.CrossesAt);
        }

        [Test]
        public void AxisDisplayUnit()
        {
            //ExStart
            //ExFor:AxisBuiltInUnit
            //ExFor:ChartAxis.DisplayUnit
            //ExFor:ChartAxis.MajorUnitIsAuto
            //ExFor:ChartAxis.MajorUnitScale
            //ExFor:ChartAxis.MinorUnitIsAuto
            //ExFor:ChartAxis.MinorUnitScale
            //ExFor:AxisDisplayUnit
            //ExFor:AxisDisplayUnit.CustomUnit
            //ExFor:AxisDisplayUnit.Unit
            //ExSummary:Shows how to manipulate the tick marks and displayed values of a chart axis.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Scatter, 450, 250);
            Chart chart = shape.Chart;

            Assert.AreEqual(1, chart.Series.Count);
            Assert.AreEqual("Y-Values", chart.Series[0].Name);

            // Set the minor tick marks of the Y-axis to point away from the plot area,
            // and the major tick marks to cross the axis.
            ChartAxis axis = chart.AxisY;
            axis.MajorTickMark = AxisTickMark.Cross;
            axis.MinorTickMark = AxisTickMark.Outside;

            // Set they Y-axis to show a major tick every 10 units, and a minor tick every 1 unit.
            axis.MajorUnit = 10;
            axis.MinorUnit = 1;

            // Set the Y-axis bounds to -10 and 20.
            // This Y-axis will now display 4 major tick marks and 27 minor tick marks.
            axis.Scaling.Minimum = new AxisBound(-10);
            axis.Scaling.Maximum = new AxisBound(20);

            // For the X-axis, set the major tick marks at every 10 units,
            // every minor tick mark at 2.5 units.
            axis = chart.AxisX;
            axis.MajorUnit = 10;
            axis.MinorUnit = 2.5;

            // Configure both types of tick marks to appear inside the graph plot area.
            axis.MajorTickMark = AxisTickMark.Inside;
            axis.MinorTickMark = AxisTickMark.Inside;

            // Set the X-axis bounds so that the X-axis spans 5 major tick marks and 12 minor tick marks.
            axis.Scaling.Minimum = new AxisBound(-10);
            axis.Scaling.Maximum = new AxisBound(30);
            axis.TickLabels.Alignment = ParagraphAlignment.Right;

            Assert.AreEqual(1, axis.TickLabels.Spacing);

            // Set the tick labels to display their value in millions.
            axis.DisplayUnit.Unit = AxisBuiltInUnit.Millions;

            // We can set a more specific value by which tick labels will display their values.
            // This statement is equivalent to the one above.
            axis.DisplayUnit.CustomUnit = 1000000;
            Assert.AreEqual(AxisBuiltInUnit.Custom, axis.DisplayUnit.Unit); //ExSkip

            doc.Save(ArtifactsDir + "Charts.AxisDisplayUnit.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.AxisDisplayUnit.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual(450.0d, shape.Width);
            Assert.AreEqual(250.0d, shape.Height);

            axis = shape.Chart.AxisX;

            Assert.AreEqual(AxisTickMark.Inside, axis.MajorTickMark);
            Assert.AreEqual(AxisTickMark.Inside, axis.MinorTickMark);
            Assert.AreEqual(10.0d, axis.MajorUnit);
            Assert.AreEqual(-10.0d, axis.Scaling.Minimum.Value);
            Assert.AreEqual(30.0d, axis.Scaling.Maximum.Value);
            Assert.AreEqual(1, axis.TickLabels.Spacing);
            Assert.AreEqual(ParagraphAlignment.Right, axis.TickLabels.Alignment);
            Assert.AreEqual(AxisBuiltInUnit.Custom, axis.DisplayUnit.Unit);
            Assert.AreEqual(1000000.0d, axis.DisplayUnit.CustomUnit);

            axis = shape.Chart.AxisY;

            Assert.AreEqual(AxisTickMark.Cross, axis.MajorTickMark);
            Assert.AreEqual(AxisTickMark.Outside, axis.MinorTickMark);
            Assert.AreEqual(10.0d, axis.MajorUnit);
            Assert.AreEqual(1.0d, axis.MinorUnit);
            Assert.AreEqual(-10.0d, axis.Scaling.Minimum.Value);
            Assert.AreEqual(20.0d, axis.Scaling.Maximum.Value);
        }

        [Test]
        public void MarkerFormatting()
        {
            //ExStart
            //ExFor:ChartMarker.Format
            //ExFor:ChartFormat.Fill
            //ExFor:ChartFormat.Stroke
            //ExFor:Stroke.ForeColor
            //ExFor:Stroke.BackColor
            //ExFor:Stroke.Visible
            //ExFor:Stroke.Transparency
            //ExFor:PresetTexture
            //ExFor:Fill.PresetTextured(PresetTexture)
            //ExSummary:Show how to set marker formatting.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Scatter, 432, 252);
            Chart chart = shape.Chart;

            // Delete default generated series.
            chart.Series.Clear();
            ChartSeries series = chart.Series.Add("AW Series 1", new[] { 0.7, 1.8, 2.6, 3.9 },
                new[] { 2.7, 3.2, 0.8, 1.7 });

            // Set marker formatting.
            series.Marker.Size = 40;
            series.Marker.Symbol = MarkerSymbol.Square;
            ChartDataPointCollection dataPoints = series.DataPoints;
            dataPoints[0].Marker.Format.Fill.PresetTextured(PresetTexture.Denim);
            dataPoints[0].Marker.Format.Stroke.ForeColor = Color.Yellow;
            dataPoints[0].Marker.Format.Stroke.BackColor = Color.Red;
            dataPoints[1].Marker.Format.Fill.PresetTextured(PresetTexture.WaterDroplets);
            dataPoints[1].Marker.Format.Stroke.ForeColor = Color.Yellow;
            dataPoints[1].Marker.Format.Stroke.Visible = false;
            dataPoints[2].Marker.Format.Fill.PresetTextured(PresetTexture.GreenMarble);
            dataPoints[2].Marker.Format.Stroke.ForeColor = Color.Yellow;
            dataPoints[3].Marker.Format.Fill.PresetTextured(PresetTexture.Oak);
            dataPoints[3].Marker.Format.Stroke.ForeColor = Color.Yellow;
            dataPoints[3].Marker.Format.Stroke.Transparency = 0.5;

            doc.Save(ArtifactsDir + "Charts.MarkerFormatting.docx");
            //ExEnd
        }

        [Test]
        public void SeriesColor()
        {
            //ExStart
            //ExFor:ChartSeries.Format
            //ExSummary:Sows how to set series color.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);

            Chart chart = shape.Chart;
            ChartSeriesCollection seriesColl = chart.Series;

            // Delete default generated series.
            seriesColl.Clear();

            // Create category names array.
            string[] categories = new[] { "Category 1", "Category 2" };

            // Adding new series. Value and category arrays must be the same size.
            ChartSeries series1 = seriesColl.Add("Series 1", categories, new double[] { 1, 2 });
            ChartSeries series2 = seriesColl.Add("Series 2", categories, new double[] { 3, 4 });
            ChartSeries series3 = seriesColl.Add("Series 3", categories, new double[] { 5, 6 });

            // Set series color.
            series1.Format.Fill.ForeColor = Color.Red;
            series2.Format.Fill.ForeColor = Color.Yellow;
            series3.Format.Fill.ForeColor = Color.Blue;

            doc.Save(ArtifactsDir + "Charts.SeriesColor.docx");
            //ExEnd
        }

        [Test]
        public void DataPointsFormatting()
        {
            //ExStart
            //ExFor:ChartDataPoint.Format
            //ExSummary:Shows how to set individual formatting for categories of a column chart.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = shape.Chart;

            // Delete default generated series.
            chart.Series.Clear();

            // Adding new series.
            ChartSeries series = chart.Series.Add("Series 1",
                new[] { "Category 1", "Category 2", "Category 3", "Category 4" },
                new double[] { 1, 2, 3, 4 });

            // Set column formatting.
            ChartDataPointCollection dataPoints = series.DataPoints;
            dataPoints[0].Format.Fill.PresetTextured(PresetTexture.Denim);
            dataPoints[1].Format.Fill.ForeColor = Color.Red;
            dataPoints[2].Format.Fill.ForeColor = Color.Yellow;
            dataPoints[3].Format.Fill.ForeColor = Color.Blue;

            doc.Save(ArtifactsDir + "Charts.DataPointsFormatting.docx");
            //ExEnd
        }

        [Test]
        public void LegendEntries()
        {
            //ExStart
            //ExFor:ChartLegendEntryCollection
            //ExFor:ChartLegend.LegendEntries
            //ExFor:ChartLegendEntry.IsHidden
            //ExSummary:Shows how to work with a legend entry for chart series.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);

            Chart chart = shape.Chart;
            ChartSeriesCollection series = chart.Series;
            series.Clear();

            string[] categories = new string[] { "AW Category 1", "AW Category 2" };

            ChartSeries series1 = series.Add("Series 1", categories, new double[] { 1, 2 });
            series.Add("Series 2", categories, new double[] { 3, 4 });
            series.Add("Series 3", categories, new double[] { 5, 6 });
            series.Add("Series 4", categories, new double[] { 0, 0 });

            ChartLegendEntryCollection legendEntries = chart.Legend.LegendEntries;
            legendEntries[3].IsHidden = true;

            doc.Save(ArtifactsDir + "Charts.LegendEntries.docx");
            //ExEnd
        }

        [Test]
        public void LegendFont()
        {
            //ExStart:LegendFont
            //GistId:470c0da51e4317baae82ad9495747fed
            //ExFor:ChartLegendEntry
            //ExFor:ChartLegendEntry.Font
            //ExFor:ChartLegend.Font
            //ExSummary:Shows how to work with a legend font.
            Document doc = new Document(MyDir + "Reporting engine template - Chart series.docx");
            Chart chart = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart;

            ChartLegend chartLegend = chart.Legend;
            // Set default font size all legend entries.
            chartLegend.Font.Size = 14;
            // Change font for specific legend entry.
            chartLegend.LegendEntries[1].Font.Italic = true;
            chartLegend.LegendEntries[1].Font.Size = 12;

            doc.Save(ArtifactsDir + "Charts.LegendFont.docx");
            //ExEnd:LegendFont
        }

        [Test]
        public void RemoveSpecificChartSeries()
        {
            //ExStart
            //ExFor:ChartSeries.SeriesType
            //ExFor:ChartSeriesType
            //ExSummary:Shows how to remove specific chart serie.
            Document doc = new Document(MyDir + "Reporting engine template - Chart series.docx");
            Chart chart = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart;

            // Remove all series of the Column type.
            for (int i = chart.Series.Count - 1; i >= 0; i--)
            {
                if (chart.Series[i].SeriesType == ChartSeriesType.Column)
                    chart.Series.RemoveAt(i);
            }

            chart.Series.Add(
                "Aspose Series",
                new string[] { "Category 1", "Category 2", "Category 3", "Category 4" },
                new double[] { 5.6, 7.1, 2.9, 8.9 });

            doc.Save(ArtifactsDir + "Charts.RemoveSpecificChartSeries.docx");
            //ExEnd
        }

        [Test]
        public void PopulateChartWithData()
        {
            //ExStart
            //ExFor:ChartXValue
            //ExFor:ChartXValue.FromDouble(Double)
            //ExFor:ChartYValue.FromDouble(Double)
            //ExFor:ChartSeries.Add(ChartXValue, ChartYValue)
            //ExSummary:Shows how to populate chart series with data.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder();

            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = shape.Chart;
            ChartSeries series1 = chart.Series[0];

            // Clear X and Y values of the first series.
            series1.ClearValues();

            // Populate the series with data.
            series1.Add(ChartXValue.FromDouble(3), ChartYValue.FromDouble(10));
            series1.Add(ChartXValue.FromDouble(5), ChartYValue.FromDouble(5));
            series1.Add(ChartXValue.FromDouble(7), ChartYValue.FromDouble(11));
            series1.Add(ChartXValue.FromDouble(9), ChartYValue.FromDouble(17));

            ChartSeries series2 = chart.Series[1];

            // Clear X and Y values of the second series.
            series2.ClearValues();

            // Populate the series with data.
            series2.Add(ChartXValue.FromDouble(2), ChartYValue.FromDouble(4));
            series2.Add(ChartXValue.FromDouble(4), ChartYValue.FromDouble(7));
            series2.Add(ChartXValue.FromDouble(6), ChartYValue.FromDouble(14));
            series2.Add(ChartXValue.FromDouble(8), ChartYValue.FromDouble(7));

            doc.Save(ArtifactsDir + "Charts.PopulateChartWithData.docx");
            //ExEnd
        }

        [Test]
        public void GetChartSeriesData()
        {
            //ExStart
            //ExFor:ChartXValueCollection
            //ExFor:ChartYValueCollection
            //ExSummary:Shows how to get chart series data.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder();

            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = shape.Chart;
            ChartSeries series = chart.Series[0];

            double minValue = double.MaxValue;
            int minValueIndex = 0;
            double maxValue = double.MinValue;
            int maxValueIndex = 0;

            for (int i = 0; i < series.YValues.Count; i++)
            {
                // Clear individual format of all data points.
                // Data points and data values are one-to-one in column charts.
                series.DataPoints[i].ClearFormat();

                // Get Y value.
                double yValue = series.YValues[i].DoubleValue;

                if (yValue < minValue)
                {
                    minValue = yValue;
                    minValueIndex = i;
                }

                if (yValue > maxValue)
                {
                    maxValue = yValue;
                    maxValueIndex = i;
                }
            }

            // Change colors of the max and min values.
            series.DataPoints[minValueIndex].Format.Fill.ForeColor = Color.Red;
            series.DataPoints[maxValueIndex].Format.Fill.ForeColor = Color.Green;

            doc.Save(ArtifactsDir + "Charts.GetChartSeriesData.docx");
            //ExEnd
        }

        [Test]
        public void ChartDataValues()
        {
            //ExStart
            //ExFor:ChartXValue.FromString(String)
            //ExFor:ChartSeries.Remove(Int32)
            //ExFor:ChartSeries.Add(ChartXValue, ChartYValue)
            //ExSummary:Shows how to add/remove chart data values.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder();

            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = shape.Chart;
            ChartSeries department1Series = chart.Series[0];
            ChartSeries department2Series = chart.Series[1];

            // Remove the first value in the both series.
            department1Series.Remove(0);
            department2Series.Remove(0);

            // Add new values to the both series.
            ChartXValue newXCategory = ChartXValue.FromString("Q1, 2023");
            department1Series.Add(newXCategory, ChartYValue.FromDouble(10.3));
            department2Series.Add(newXCategory, ChartYValue.FromDouble(5.7));

            doc.Save(ArtifactsDir + "Charts.ChartDataValues.docx");
            //ExEnd
        }

        [Test]
        public void FormatDataLables()
        {
            //ExStart
            //ExFor:ChartDataLabelCollection.Format
            //ExFor:ChartFormat.ShapeType
            //ExFor:ChartShapeType
            //ExSummary:Shows how to set fill, stroke and callout formatting for chart data labels.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = shape.Chart;

            // Delete default generated series.
            chart.Series.Clear();

            // Add new series.
            ChartSeries series = chart.Series.Add("AW Series 1",
                new string[] { "AW Category 1", "AW Category 2", "AW Category 3", "AW Category 4" },
                new double[] { 100, 200, 300, 400 });

            // Show data labels.
            series.HasDataLabels = true;
            series.DataLabels.ShowValue = true;

            // Format data labels as callouts.
            ChartFormat format = series.DataLabels.Format;
            format.ShapeType = ChartShapeType.WedgeRectCallout;
            format.Stroke.Color = Color.DarkGreen;
            format.Fill.Solid(Color.Green);
            series.DataLabels.Font.Color = Color.Yellow;

            // Change fill and stroke of an individual data label.
            ChartFormat labelFormat = series.DataLabels[0].Format;
            labelFormat.Stroke.Color = Color.DarkBlue;
            labelFormat.Fill.Solid(Color.Blue);

            doc.Save(ArtifactsDir + "Charts.FormatDataLables.docx");
            //ExEnd
        }

        [Test]
        public void ChartAxisTitle()
        {
            //ExStart:ChartAxisTitle
            //GistId:3428e84add5beb0d46a8face6e5fc858
            //ExFor:ChartAxisTitle
            //ExFor:ChartAxisTitle.Text
            //ExFor:ChartAxisTitle.Show
            //ExFor:ChartAxisTitle.Overlay
            //ExFor:ChartAxisTitle.Font
            //ExSummary:Shows how to set chart axis title.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);

            Chart chart = shape.Chart;
            ChartSeriesCollection seriesColl = chart.Series;
            // Delete default generated series.
            seriesColl.Clear();

            seriesColl.Add("AW Series 1", new string[] { "AW Category 1", "AW Category 2" }, new double[] { 1, 2 });

            ChartAxisTitle chartAxisXTitle = chart.AxisX.Title;
            chartAxisXTitle.Text = "Categories";
            chartAxisXTitle.Show = true;
            ChartAxisTitle chartAxisYTitle = chart.AxisY.Title;
            chartAxisYTitle.Text = "Values";
            chartAxisYTitle.Show = true;
            chartAxisYTitle.Overlay = true;
            chartAxisYTitle.Font.Size = 12;
            chartAxisYTitle.Font.Color = Color.Blue;

            doc.Save(ArtifactsDir + "Charts.ChartAxisTitle.docx");
            //ExEnd:ChartAxisTitle
        }

        [TestCase(new[] { 1, 2, double.NaN, 4, 5, 6 })]
        [TestCase(new[] { double.NaN, 4, 5, double.NaN, 7, 8 })]
        [TestCase(new[] { double.NaN, double.NaN, double.NaN, double.NaN, double.NaN, 9 })]
        [TestCase(new[] { double.NaN, 4, 5, double.NaN, double.NaN }, typeof(ArgumentException))]
        [TestCase(new[] { double.NaN, double.NaN, double.NaN, double.NaN, double.NaN }, typeof(ArgumentException))]
        public void DataArraysWrongSize(double[] seriesValue, Type exception = null)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Line, 500, 300);
            ChartSeriesCollection seriesColl = shape.Chart.Series;
            seriesColl.Clear();

            string[] categories = { "Word", null, "Excel", "GoogleDocs", "Note", null };
            if (exception is null)
                seriesColl.Add("AW Series", categories, seriesValue);
            else
                Assert.Throws(exception, () => seriesColl.Add("AW Series", categories, seriesValue));
        }

        [Test]
        public void CopyDataPointFormat()
        {
            //ExStart:CopyDataPointFormat
            //GistId:3428e84add5beb0d46a8face6e5fc858
            //ExFor:ChartSeries.CopyFormatFrom(int)
            //ExFor:ChartDataPointCollection.HasDefaultFormat(int)
            //ExFor:ChartDataPointCollection.CopyFormat(int, int)
            //ExSummary:Shows how to copy data point format.
            Document doc = new Document(MyDir + "DataPoint format.docx");

            // Get the chart and series to update format.
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            ChartSeries series = shape.Chart.Series[0];
            ChartDataPointCollection dataPoints = series.DataPoints;

            Assert.IsTrue(dataPoints.HasDefaultFormat(0));
            Assert.IsFalse(dataPoints.HasDefaultFormat(1));

            // Copy format of the data point with index 1 to the data point with index 2
            // so that the data point 2 looks the same as the data point 1.
            dataPoints.CopyFormat(0, 1);

            Assert.IsTrue(dataPoints.HasDefaultFormat(0));
            Assert.IsTrue(dataPoints.HasDefaultFormat(1));

            // Copy format of the data point with index 0 to the series defaults so that all data points
            // in the series that have the default format look the same as the data point 0.
            series.CopyFormatFrom(1);

            Assert.IsTrue(dataPoints.HasDefaultFormat(0));
            Assert.IsTrue(dataPoints.HasDefaultFormat(1));

            doc.Save(ArtifactsDir + "Charts.CopyDataPointFormat.docx");
            //ExEnd:CopyDataPointFormat
        }

        [Test]
        public void ResetDataPointFill()
        {
            //ExStart:ResetDataPointFill
            //GistId:3428e84add5beb0d46a8face6e5fc858
            //ExFor:ChartFormat.IsDefined
            //ExFor:ChartFormat.SetDefaultFill
            //ExSummary:Shows how to reset the fill to the default value defined in the series.
            Document doc = new Document(MyDir + "DataPoint format.docx");

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            ChartSeries series = shape.Chart.Series[0];
            ChartDataPoint dataPoint = series.DataPoints[1];

            Assert.IsTrue(dataPoint.Format.IsDefined);

            dataPoint.Format.SetDefaultFill();

            doc.Save(ArtifactsDir + "Charts.ResetDataPointFill.docx");
            //ExEnd:ResetDataPointFill
        }

        [Test]
        public void DataTable()
        {
            //ExStart:DataTable
            //GistId:a775441ecb396eea917a2717cb9e8f8f
            //ExFor:Chart.DataTable
            //ExFor:ChartDataTable
            //ExFor:ChartDataTable.Show
            //ExSummary:Shows how to show data table with chart series data.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = shape.Chart;
            
            ChartSeriesCollection series = chart.Series;
            series.Clear();
            double[] xValues = new double[] { 2020, 2021, 2022, 2023 };
            series.Add("Series1", xValues, new double[] { 5, 11, 2, 7 });
            series.Add("Series2", xValues, new double[] { 6, 5.5, 7, 7.8 });
            series.Add("Series3", xValues, new double[] { 10, 8, 7, 9 });

            ChartDataTable dataTable = chart.DataTable;
            dataTable.Show = true;

            dataTable.HasLegendKeys = false;
            dataTable.HasHorizontalBorder = false;
            dataTable.HasVerticalBorder = false;

            dataTable.Font.Italic = true;
            dataTable.Format.Stroke.Weight = 1;
            dataTable.Format.Stroke.DashStyle = DashStyle.ShortDot;
            dataTable.Format.Stroke.Color = Color.DarkBlue;

            doc.Save(ArtifactsDir + "Charts.DataTable.docx");
            //ExEnd:DataTable
        }

        [Test]
        public void ChartFormat()
        {
            //ExStart:ChartFormat
            //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
            //ExFor:ChartFormat
            //ExFor:Chart.Format
            //ExFor:ChartTitle.Format
            //ExFor:ChartAxisTitle.Format
            //ExFor:ChartLegend.Format
            //ExFor:Fill.Solid(Color)
            //ExSummary:Shows how to use chart formating.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = shape.Chart;

            // Delete series generated by default.
            ChartSeriesCollection series = chart.Series;
            series.Clear();

            string[] categories = new string[] { "Category 1", "Category 2" };
            series.Add("Series 1", categories, new double[] { 1, 2 });
            series.Add("Series 2", categories, new double[] { 3, 4 });

            // Format chart background.
            chart.Format.Fill.Solid(Color.DarkSlateGray);

            // Hide axis tick labels.
            chart.AxisX.TickLabels.Position = AxisTickLabelPosition.None;
            chart.AxisY.TickLabels.Position = AxisTickLabelPosition.None;

            // Format chart title.
            chart.Title.Format.Fill.Solid(Color.LightGoldenrodYellow);

            // Format axis title.
            chart.AxisX.Title.Show = true;
            chart.AxisX.Title.Format.Fill.Solid(Color.LightGoldenrodYellow);

            // Format legend.
            chart.Legend.Format.Fill.Solid(Color.LightGoldenrodYellow);

            doc.Save(ArtifactsDir + "Charts.ChartFormat.docx");
            //ExEnd:ChartFormat

            doc = new Document(ArtifactsDir + "Charts.ChartFormat.docx");

            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            chart = shape.Chart;

            Assert.AreEqual(Color.DarkSlateGray.ToArgb(), chart.Format.Fill.Color.ToArgb());
            Assert.AreEqual(Color.LightGoldenrodYellow.ToArgb(), chart.Title.Format.Fill.Color.ToArgb());
            Assert.AreEqual(Color.LightGoldenrodYellow.ToArgb(), chart.AxisX.Title.Format.Fill.Color.ToArgb());
            Assert.AreEqual(Color.LightGoldenrodYellow.ToArgb(), chart.Legend.Format.Fill.Color.ToArgb());
        }

        [Test]
        public void SecondaryAxis()
        {
            //ExStart:SecondaryAxis
            //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
            //ExFor:ChartSeriesGroup
            //ExFor:ChartSeriesGroup.AxisGroup
            //ExFor:ChartSeriesGroup.AxisX
            //ExFor:ChartSeriesGroup.AxisY
            //ExFor:ChartSeriesGroup.Series
            //ExFor:ChartSeriesGroupCollection
            //ExFor:ChartSeriesGroupCollection.Add(ChartSeriesType)
            //ExFor:AxisGroup
            //ExSummary:Shows how to work with the secondary axis of chart.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Line, 450, 250);
            Chart chart = shape.Chart;
            ChartSeriesCollection series = chart.Series;

            // Delete default generated series.
            series.Clear();

            string[] categories = new string[] { "Category 1", "Category 2", "Category 3" };
            series.Add("Series 1 of primary series group", categories, new double[] { 2, 3, 4 });
            series.Add("Series 2 of primary series group", categories, new double[] { 5, 2, 3 });

            // Create an additional series group, also of the line type.
            ChartSeriesGroup newSeriesGroup = chart.SeriesGroups.Add(ChartSeriesType.Line);
            // Specify the use of secondary axes for the new series group.
            newSeriesGroup.AxisGroup = AxisGroup.Secondary;
            // Hide the secondary X axis.
            newSeriesGroup.AxisX.Hidden = true;
            // Define title of the secondary Y axis.
            newSeriesGroup.AxisY.Title.Show = true;
            newSeriesGroup.AxisY.Title.Text = "Secondary Y axis";

            // Add a series to the new series group.
            ChartSeries series3 =
                newSeriesGroup.Series.Add("Series of secondary series group", categories, new double[] { 13, 11, 16 });
            series3.Format.Stroke.Weight = 3.5;

            doc.Save(ArtifactsDir + "Charts.SecondaryAxis.docx");
            //ExEnd:SecondaryAxis
        }

        [Test]
        public void ConfigureGapOverlap()
        {
            //ExStart:ConfigureGapOverlap
            //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
            //ExFor:Chart.SeriesGroups
            //ExFor:ChartSeriesGroup.GapWidth
            //ExFor:ChartSeriesGroup.Overlap
            //ExSummary:Show how to configure gap width and overlap.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Column, 450, 250);
            ChartSeriesGroup seriesGroup = shape.Chart.SeriesGroups[0];

            // Set column gap width and overlap.
            seriesGroup.GapWidth = 450;
            seriesGroup.Overlap = -75;

            doc.Save(ArtifactsDir + "Charts.ConfigureGapOverlap.docx");
            //ExEnd:ConfigureGapOverlap
        }

        [Test]
        public void BubbleScale()
        {
            //ExStart:BubbleScale
            //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
            //ExFor:ChartSeriesGroup.BubbleScale
            //ExSummary:Show how to set size of the bubbles.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a bubble 3D chart.
            Shape shape = builder.InsertChart(ChartType.Bubble3D, 450, 250);
            ChartSeriesGroup seriesGroup = shape.Chart.SeriesGroups[0];

            // Set bubble scale to 200%.
            seriesGroup.BubbleScale = 200;

            doc.Save(ArtifactsDir + "Charts.BubbleScale.docx");
            //ExEnd:BubbleScale
        }

        [Test]
        public void RemoveSecondaryAxis()
        {
            //ExStart:RemoveSecondaryAxis
            //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
            //ExFor:ChartSeriesGroupCollection.Count
            //ExFor:ChartSeriesGroupCollection.Item(Int32)
            //ExFor:ChartSeriesGroupCollection.RemoveAt(Int32)
            //ExSummary:Show how to remove secondary axis.
            Document doc = new Document(MyDir + "Combo chart.docx");

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Chart chart = shape.Chart;
            ChartSeriesGroupCollection seriesGroups = chart.SeriesGroups;

            // Find secondary axis and remove from the collection.
            for (int i = 0; i < seriesGroups.Count; i++)
                if (seriesGroups[i].AxisGroup == AxisGroup.Secondary)
                    seriesGroups.RemoveAt(i);
            //ExEnd:RemoveSecondaryAxis
        }

        [Test]
        public void TreemapChart()
        {
            //ExStart:TreemapChart
            //GistId:65919861586e42e24f61a3ccb65f8f4e
            //ExFor:ChartSeriesCollection.Add(String, ChartMultilevelValue[], double[])
            //ExFor:ChartMultilevelValue
            //ExFor:ChartMultilevelValue.#ctor(String, String, String)
            //ExFor:ChartMultilevelValue.#ctor(String, String)
            //ExFor:ChartMultilevelValue.#ctor(String)
            //ExSummary:Shows how to create treemap chart.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Treemap chart.
            Shape shape = builder.InsertChart(ChartType.Treemap, 450, 280);
            Chart chart = shape.Chart;
            chart.Title.Text = "World Population";

            // Delete default generated series.
            chart.Series.Clear();

            // Add a series.
            ChartSeries series = chart.Series.Add(
                "Population by Region",
                new ChartMultilevelValue[]
                {
                    new ChartMultilevelValue("Asia", "China"),
                    new ChartMultilevelValue("Asia", "India"),
                    new ChartMultilevelValue("Asia", "Indonesia"),
                    new ChartMultilevelValue("Asia", "Pakistan"),
                    new ChartMultilevelValue("Asia", "Bangladesh"),
                    new ChartMultilevelValue("Asia", "Japan"),
                    new ChartMultilevelValue("Asia", "Philippines"),
                    new ChartMultilevelValue("Asia", "Other"),
                    new ChartMultilevelValue("Africa", "Nigeria"),
                    new ChartMultilevelValue("Africa", "Ethiopia"),
                    new ChartMultilevelValue("Africa", "Egypt"),
                    new ChartMultilevelValue("Africa", "Other"),
                    new ChartMultilevelValue("Europe", "Russia"),
                    new ChartMultilevelValue("Europe", "Germany"),
                    new ChartMultilevelValue("Europe", "Other"),
                    new ChartMultilevelValue("Latin America", "Brazil"),
                    new ChartMultilevelValue("Latin America", "Mexico"),
                    new ChartMultilevelValue("Latin America", "Other"),
                    new ChartMultilevelValue("Northern America", "United States", "Other"),
                    new ChartMultilevelValue("Northern America", "Other"),
                    new ChartMultilevelValue("Oceania")
                },
                new double[]
                {
                    1409670000, 1400744000, 279118866, 241499431, 169828911, 123930000, 112892781, 764000000,
                    223800000, 107334000, 105914499, 903000000,
                    146150789, 84607016, 516000000,
                    203080756, 129713690, 310000000,
                    335893238, 35000000,
                    42000000
                });

            // Show data labels.
            series.HasDataLabels = true;
            series.DataLabels.ShowValue = true;
            series.DataLabels.ShowCategoryName = true;
            string thousandSeparator = CultureInfo.CurrentCulture.NumberFormat.CurrencyGroupSeparator;
            series.DataLabels.NumberFormat.FormatCode = $"#{thousandSeparator}0";

            doc.Save(ArtifactsDir + "Charts.Treemap.docx");
            //ExEnd:TreemapChart
        }

        [Test]
        public void SunburstChart()
        {
            //ExStart:SunburstChart
            //GistId:65919861586e42e24f61a3ccb65f8f4e
            //ExFor:ChartSeriesCollection.Add(String, ChartMultilevelValue[], double[])
            //ExSummary:Shows how to create sunburst chart.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Sunburst chart.
            Shape shape = builder.InsertChart(ChartType.Sunburst, 450, 450);
            Chart chart = shape.Chart;
            chart.Title.Text = "Sales";

            // Delete default generated series.
            chart.Series.Clear();

            // Add a series.
            ChartSeries series = chart.Series.Add(
                "Sales",
                new ChartMultilevelValue[]
                {
                    new ChartMultilevelValue("Sales - Europe", "UK", "London Dep."),
                    new ChartMultilevelValue("Sales - Europe", "UK", "Liverpool Dep."),
                    new ChartMultilevelValue("Sales - Europe", "UK", "Manchester Dep."),
                    new ChartMultilevelValue("Sales - Europe", "France", "Paris Dep."),
                    new ChartMultilevelValue("Sales - Europe", "France", "Lyon Dep."),
                    new ChartMultilevelValue("Sales - NA", "USA", "Denver Dep."),
                    new ChartMultilevelValue("Sales - NA", "USA", "Seattle Dep."),
                    new ChartMultilevelValue("Sales - NA", "USA", "Detroit Dep."),
                    new ChartMultilevelValue("Sales - NA", "USA", "Houston Dep."),
                    new ChartMultilevelValue("Sales - NA", "Canada", "Toronto Dep."),
                    new ChartMultilevelValue("Sales - NA", "Canada", "Montreal Dep."),
                    new ChartMultilevelValue("Sales - Oceania", "Australia", "Sydney Dep."),
                    new ChartMultilevelValue("Sales - Oceania", "New Zealand", "Auckland Dep.")
                },
                new double[] { 1236, 851, 536, 468, 179, 527, 799, 1148, 921, 457, 482, 761, 694 });

            // Show data labels.
            series.HasDataLabels = true;
            series.DataLabels.ShowValue = false;
            series.DataLabels.ShowCategoryName = true;

            doc.Save(ArtifactsDir + "Charts.Sunburst.docx");
            //ExEnd:SunburstChart
        }

        [Test]
        public void HistogramChart()
        {
            //ExStart:HistogramChart
            //GistId:65919861586e42e24f61a3ccb65f8f4e
            //ExFor:ChartSeriesCollection.Add(String, double[])
            //ExSummary:Shows how to create histogram chart.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Histogram chart.
            Shape shape = builder.InsertChart(ChartType.Histogram, 450, 450);
            Chart chart = shape.Chart;
            chart.Title.Text = "Avg Temperature since 1991";

            // Delete default generated series.
            chart.Series.Clear();

            // Add a series.
            chart.Series.Add(
                "Avg Temperature",
                new double[]
                {
                    51.8, 53.6, 50.3, 54.7, 53.9, 54.3, 53.4, 52.9, 53.3, 53.7, 53.8, 52.0, 55.0, 52.1, 53.4,
                    53.8, 53.8, 51.9, 52.1, 52.7, 51.8, 56.6, 53.3, 55.6, 56.3, 56.2, 56.1, 56.2, 53.6, 55.7,
                    56.3, 55.9, 55.6
                });

            doc.Save(ArtifactsDir + "Charts.Histogram.docx");
            //ExEnd:HistogramChart
        }

        [Test]
        public void ParetoChart()
        {
            //ExStart:ParetoChart
            //GistId:65919861586e42e24f61a3ccb65f8f4e
            //ExFor:ChartSeriesCollection.Add(String, String[], double[])
            //ExSummary:Shows how to create pareto chart.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Pareto chart.
            Shape shape = builder.InsertChart(ChartType.Pareto, 450, 450);
            Chart chart = shape.Chart;
            chart.Title.Text = "Best-Selling Car";

            // Delete default generated series.
            chart.Series.Clear();

            // Add a series.
            chart.Series.Add(
                "Best-Selling Car",
                new string[] { "Tesla Model Y", "Toyota Corolla", "Toyota RAV4", "Ford F-Series", "Honda CR-V" },
                new double[] { 1.43, 0.91, 1.17, 0.98, 0.85 });

            doc.Save(ArtifactsDir + "Charts.Pareto.docx");
            //ExEnd:ParetoChart
        }

        [Test]
        public void BoxAndWhiskerChart()
        {
            //ExStart:BoxAndWhiskerChart
            //GistId:65919861586e42e24f61a3ccb65f8f4e
            //ExFor:ChartSeriesCollection.Add(String, String[], double[])
            //ExSummary:Shows how to create box and whisker chart.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Box & Whisker chart.
            Shape shape = builder.InsertChart(ChartType.BoxAndWhisker, 450, 450);
            Chart chart = shape.Chart;
            chart.Title.Text = "Points by Years";

            // Delete default generated series.
            chart.Series.Clear();

            // Add a series.
            ChartSeries series = chart.Series.Add(
                "Points by Years",
                new string[]
                {
                    "WC", "WC", "WC", "WC", "WC", "WC", "WC", "WC", "WC", "WC",
                    "NR", "NR", "NR", "NR", "NR", "NR", "NR", "NR", "NR", "NR",
                    "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA"
                },
                new double[]
                {
                    91, 80, 100, 77, 90, 104, 105, 118, 120, 101,
                    114, 107, 110, 60, 79, 78, 77, 102, 101, 113,
                    94, 93, 84, 71, 80, 103, 80, 94, 100, 101
                });

            // Show data labels.
            series.HasDataLabels = true;

            doc.Save(ArtifactsDir + "Charts.BoxAndWhisker.docx");
            //ExEnd:BoxAndWhiskerChart
        }

        [Test]
        public void WaterfallChart()
        {
            //ExStart:WaterfallChart
            //GistId:65919861586e42e24f61a3ccb65f8f4e
            //ExFor:ChartSeriesCollection.Add(String, String[], double[], bool[])
            //ExSummary:Shows how to create waterfall chart.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Waterfall chart.
            Shape shape = builder.InsertChart(ChartType.Waterfall, 450, 450);
            Chart chart = shape.Chart;
            chart.Title.Text = "New Zealand GDP";

            // Delete default generated series.
            chart.Series.Clear();

            // Add a series.
            ChartSeries series = chart.Series.Add(
                "New Zealand GDP",
                new string[] { "2018", "2019 growth", "2020 growth", "2020", "2021 growth", "2022 growth", "2022" },
                new double[] { 100, 0.57, -0.25, 100.32, 20.22, -2.92, 117.62 },
                new bool[] { true, false, false, true, false, false, true });

            // Show data labels.
            series.HasDataLabels = true;

            doc.Save(ArtifactsDir + "Charts.Waterfall.docx");
            //ExEnd:WaterfallChart
        }

        [Test]
        public void FunnelChart()
        {
            //ExStart:FunnelChart
            //GistId:65919861586e42e24f61a3ccb65f8f4e
            //ExFor:ChartSeriesCollection.Add(String, String[], double[])
            //ExSummary:Shows how to create funnel chart.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Funnel chart.
            Shape shape = builder.InsertChart(ChartType.Funnel, 450, 450);
            Chart chart = shape.Chart;
            chart.Title.Text = "Population by Age Group";

            // Delete default generated series.
            chart.Series.Clear();

            // Add a series.
            ChartSeries series = chart.Series.Add(
                "Population by Age Group",
                new string[] { "0-9", "10-19", "20-29", "30-39", "40-49", "50-59", "60-69", "70-79", "80-89", "90-" },
                new double[] { 0.121, 0.128, 0.132, 0.146, 0.124, 0.124, 0.111, 0.075, 0.032, 0.007 });

            // Show data labels.
            series.HasDataLabels = true;
            string decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator;
            series.DataLabels.NumberFormat.FormatCode = $"0{decimalSeparator}0%";

            doc.Save(ArtifactsDir + "Charts.Funnel.docx");
            //ExEnd:FunnelChart
        }

        [Test]
        public void LabelOrientationRotation()
        {
            //ExStart:LabelOrientationRotation
            //GistId:ac8ba4eb35f3fbb8066b48c999da63b0
            //ExFor:ChartDataLabelCollection.Orientation
            //ExFor:ChartDataLabelCollection.Rotation
            //ExFor:ChartDataLabel.Rotation
            //ExFor:ChartDataLabel.Orientation
            //ExFor:ShapeTextOrientation
            //ExSummary:Shows how to change orientation and rotation for data labels.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            ChartSeries series = shape.Chart.Series[0];
            ChartDataLabelCollection dataLabels = series.DataLabels;

            // Show data labels.
            series.HasDataLabels = true;
            dataLabels.ShowValue = true;
            dataLabels.ShowCategoryName = true;

            // Define data label shape.
            dataLabels.Format.ShapeType = ChartShapeType.UpArrow;
            dataLabels.Format.Stroke.Fill.Solid(Color.DarkBlue);

            // Set data label orientation and rotation for the entire series.
            dataLabels.Orientation = ShapeTextOrientation.VerticalFarEast;
            dataLabels.Rotation = -45;

            // Change orientation and rotation of the first data label.
            dataLabels[0].Orientation = ShapeTextOrientation.Horizontal;
            dataLabels[0].Rotation = 45;

            doc.Save(ArtifactsDir + "Charts.LabelOrientationRotation.docx");
            //ExEnd:LabelOrientationRotation
        }

        [Test]
        public void TickLabelsOrientationRotation()
        {
            //ExStart:TickLabelsOrientationRotation
            //GistId:708ce40a68fac5003d46f6b4acfd5ff1
            //ExFor:AxisTickLabels.Rotation
            //ExFor:AxisTickLabels.Orientation
            //ExSummary:Shows how to change orientation and rotation for axis tick labels.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a column chart.
            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            AxisTickLabels xTickLabels = shape.Chart.AxisX.TickLabels;
            AxisTickLabels yTickLabels = shape.Chart.AxisY.TickLabels;

            // Set axis tick label orientation and rotation.
            xTickLabels.Orientation = ShapeTextOrientation.VerticalFarEast;
            xTickLabels.Rotation = -30;
            yTickLabels.Orientation = ShapeTextOrientation.Horizontal;
            yTickLabels.Rotation = 45;

            doc.Save(ArtifactsDir + "Charts.TickLabelsOrientationRotation.docx");
            //ExEnd:TickLabelsOrientationRotation
        }

        [Test]
        public void DoughnutChart()
        {
            //ExStart:DoughnutChart
            //GistId:bb594993b5fe48692541e16f4d354ac2
            //ExFor:ChartSeriesGroup.DoughnutHoleSize
            //ExFor:ChartSeriesGroup.FirstSliceAngle
            //ExSummary:Shows how to create and format Doughnut chart.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Doughnut, 400, 400);
            Chart chart = shape.Chart;
            // Delete the default generated series.
            chart.Series.Clear();

            string[] categories = new string[] { "Category 1", "Category 2", "Category 3" };
            chart.Series.Add("Series 1", categories, new double[] { 4, 2, 5 });

            // Format the Doughnut chart.
            ChartSeriesGroup seriesGroup = chart.SeriesGroups[0];
            seriesGroup.DoughnutHoleSize = 10;
            seriesGroup.FirstSliceAngle = 270;

            doc.Save(ArtifactsDir + "Charts.DoughnutChart.docx");
            //ExEnd:DoughnutChart
        }

        [Test]
        public void PieOfPieChart()
        {
            //ExStart:PieOfPieChart
            //GistId:bb594993b5fe48692541e16f4d354ac2
            //ExFor:ChartSeriesGroup.SecondSectionSize
            //ExSummary:Shows how to create and format pie of Pie chart.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.PieOfPie, 440, 300);
            Chart chart = shape.Chart;
            // Delete the default generated series.
            chart.Series.Clear();

            string[] categories = new string[] { "Category 1", "Category 2", "Category 3", "Category 4" };
            chart.Series.Add("Series 1", categories, new double[] { 11, 8, 4, 3 });

            // Format the Pie of Pie chart.
            ChartSeriesGroup seriesGroup = chart.SeriesGroups[0];
            seriesGroup.GapWidth = 10;
            seriesGroup.SecondSectionSize = 77;

            doc.Save(ArtifactsDir + "Charts.PieOfPieChart.docx");
            //ExEnd:PieOfPieChart
        }

        [Test]
        public void FormatCode()
        {
            //ExStart:FormatCode
            //ReleaseVersion:24.11
            //ExFor:ChartXValueCollection.FormatCode
            //ExFor:ChartYValueCollection.FormatCode
            //ExFor:BubbleSizeCollection.FormatCode
            //ExSummary:Shows how to work with the format code of the chart data.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a Bubble chart.
            Shape shape = builder.InsertChart(ChartType.Bubble, 432, 252);
            Chart chart = shape.Chart;

            // Delete default generated series.
            chart.Series.Clear();

            ChartSeries series = chart.Series.Add(
                "Series1",
                new double[] { 1, 1.9, 2.45, 3 },
                new double[] { 1, -0.9, 1.82, 0 },
                new double[] { 2, 1.1, 2.95, 2 });

            // Show data labels.
            series.HasDataLabels = true;
            series.DataLabels.ShowCategoryName = true;
            series.DataLabels.ShowValue = true;
            series.DataLabels.ShowBubbleSize = true;

            // Set data format codes.
            series.XValues.FormatCode = "#,##0.0#";
            series.YValues.FormatCode = "#,##0.0#;[Red]\\-#,##0.0#";
            series.BubbleSizes.FormatCode = "#,##0.0#";

            doc.Save(ArtifactsDir + "Charts.FormatCode.docx");
            //ExEnd:FormatCode

            doc = new Document(ArtifactsDir + "Charts.FormatCode.docx");
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            chart = shape.Chart;

            ChartSeriesCollection seriesCollection = chart.Series;
            foreach (ChartSeries seriesProperties in seriesCollection)
            {
                Assert.AreEqual("#,##0.0#", seriesProperties.XValues.FormatCode);
                Assert.AreEqual("#,##0.0#;[Red]\\-#,##0.0#", seriesProperties.YValues.FormatCode);
                Assert.AreEqual("#,##0.0#", seriesProperties.BubbleSizes.FormatCode);
            }
        }
    }
}
