// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExCharts : ApiExampleBase
    {
        [Test]
        public void ChartTitle()
        {
            //ExStart
            //ExFor:Chart
            //ExFor:Chart.Title
            //ExFor:ChartTitle
            //ExFor:ChartTitle.Overlay
            //ExFor:ChartTitle.Show
            //ExFor:ChartTitle.Text
            //ExSummary:Shows how to insert a chart and set a title.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a chart shape with a document builder and get its chart.
            Shape chartShape = builder.InsertChart(ChartType.Bar, 400, 300);
            Chart chart = chartShape.Chart;

            // Set the title text, which appears at the top center of the chart.
            // Then, set it to be shown, and also allow other graph elements to overlap it.
            ChartTitle title = chart.Title;
            title.Text = "My Chart";
            title.Show = true;
            title.Overlay = true;

            doc.Save(ArtifactsDir + "Charts.ChartTitle.docx");
            //ExEnd

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
            //ExFor:ChartNumberFormat.FormatCode
            //ExSummary:Shows how to enable and configure data labels for a chart series.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a chart, clear its demo data, and set a name.
            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);
            Chart chart = shape.Chart;
            chart.Series.Clear();
            chart.Title.Text = "Monthly sales report";
            
            // Insert a custom chart series.
            ChartSeries series = chart.Series.Add("Revenue", 
                new[] { "January", "February", "March" }, new[] { 25.611d, 21.439d, 33.750d });

            // Enable data labels and set a custom number format with which to display decimal series values.
            series.HasDataLabels = true;
            ChartDataLabelCollection dataLabels = series.DataLabels;
            dataLabels.ShowValue = true;
            dataLabels.NumberFormat.FormatCode = "\"US$\" #,##0.000\"M\"";

            doc.Save(ArtifactsDir + "Charts.DataLabelNumberFormat.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.DataLabelNumberFormat.docx");
            series = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart.Series[0];

            Assert.True(series.HasDataLabels);
            Assert.True(series.DataLabels.ShowValue);
            Assert.AreEqual("\"US$\" #,##0.000\"M\"", series.DataLabels.NumberFormat.FormatCode);
        }

        [Test]
        public void DataArraysWrongSize()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add chart with default data
            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);
            Chart chart = shape.Chart;

            ChartSeriesCollection seriesColl = chart.Series;
            seriesColl.Clear();

            // Create category names array, second category will be null
            string[] categories = { "Cat1", null, "Cat3", "Cat4", "Cat5", null };

            // Adding new series with empty (double.NaN) values
            seriesColl.Add("AW Series 1", categories, new double[] { 1, 2, double.NaN, 4, 5, 6 });
            seriesColl.Add("AW Series 2", categories, new double[] { 2, 3, double.NaN, 5, 6, 7 });

            Assert.That(
                () => seriesColl.Add("AW Series 3", categories, new[] { double.NaN, 4, 5, double.NaN, double.NaN }),
                Throws.TypeOf<ArgumentException>());
            Assert.That(
                () => seriesColl.Add("AW Series 4", categories,
                    new[] { double.NaN, double.NaN, double.NaN, double.NaN, double.NaN }),
                Throws.TypeOf<ArgumentException>());
        }

        [Test]
        public void EmptyValuesInChartData()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add chart with default data
            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);
            Chart chart = shape.Chart;

            ChartSeriesCollection seriesColl = chart.Series;
            seriesColl.Clear();

            // Create category names array, second category will be null
            string[] categories = { "Cat1", null, "Cat3", "Cat4", "Cat5", null };

            // Adding new series with empty (double.NaN) values
            seriesColl.Add("AW Series 1", categories, new[] { 1, 2, double.NaN, 4, 5, 6 });
            seriesColl.Add("AW Series 2", categories, new[] { 2, 3, double.NaN, 5, 6, 7 });
            seriesColl.Add("AW Series 3", categories, new[] { double.NaN, 4, 5, double.NaN, 7, 8 });
            seriesColl.Add("AW Series 4", categories,
                new[] { double.NaN, double.NaN, double.NaN, double.NaN, double.NaN, 9 });

            doc.Save(ArtifactsDir + "Charts.EmptyValuesInChartData.docx");
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
            //ExFor:ChartAxis.TickLabelOffset
            //ExFor:ChartAxis.TickLabelPosition
            //ExFor:ChartAxis.TickLabelSpacingIsAuto
            //ExFor:ChartAxis.TickMarkSpacing
            //ExFor:Charts.AxisCategoryType
            //ExFor:Charts.AxisCrosses
            //ExFor:Charts.Chart.AxisX
            //ExFor:Charts.Chart.AxisY
            //ExFor:Charts.Chart.AxisZ
            //ExSummary:Shows how to insert a chart and modify the appearance of its axes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a column chart, clear the demo data, and then add a custom data series.
            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = shape.Chart;
            chart.Series.Clear();
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
            xAxis.TickLabelOffset = 50;
            xAxis.TickLabelPosition = AxisTickLabelPosition.Low;
            xAxis.TickLabelSpacingIsAuto = false;
            xAxis.TickMarkSpacing = 1;

            ChartAxis yAxis = chart.AxisY;
            yAxis.CategoryType = AxisCategoryType.Automatic;
            yAxis.Crosses = AxisCrosses.Maximum;
            yAxis.ReverseOrder = true;
            yAxis.MajorTickMark = AxisTickMark.Inside;
            yAxis.MinorTickMark = AxisTickMark.Cross;
            yAxis.MajorUnit = 100.0d;
            yAxis.MinorUnit = 20.0d;
            yAxis.TickLabelPosition = AxisTickLabelPosition.NextToAxis;

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
            Assert.AreEqual(50, chart.AxisX.TickLabelOffset);
            Assert.AreEqual(AxisTickLabelPosition.Low, chart.AxisX.TickLabelPosition);
            Assert.False(chart.AxisX.TickLabelSpacingIsAuto);
            Assert.AreEqual(1, chart.AxisX.TickMarkSpacing);

            Assert.AreEqual(AxisCategoryType.Category, chart.AxisY.CategoryType);
            Assert.AreEqual(AxisCrosses.Maximum, chart.AxisY.Crosses);
            Assert.True(chart.AxisY.ReverseOrder);
            Assert.AreEqual(AxisTickMark.Inside, chart.AxisY.MajorTickMark);
            Assert.AreEqual(AxisTickMark.Cross, chart.AxisY.MinorTickMark);
            Assert.AreEqual(100.0d, chart.AxisY.MajorUnit);
            Assert.AreEqual(20.0d, chart.AxisY.MinorUnit);
            Assert.AreEqual(AxisTickLabelPosition.NextToAxis, chart.AxisY.TickLabelPosition);
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
            //ExFor:Charts.AxisTickMark
            //ExFor:Charts.AxisTickLabelPosition
            //ExFor:Charts.AxisTimeUnit
            //ExFor:Charts.ChartAxis.BaseTimeUnit
            //ExSummary:Shows how to insert chart with date/time values.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a line chart, clear its demo data,
            // and then add a custom series containing date/time values for the X-axis, and decimal values for the Y-axis.
            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);
            Chart chart = shape.Chart;
            chart.Series.Clear();
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

            // Define Y-axis properties for decimal values.
            ChartAxis yAxis = chart.AxisY;
            yAxis.TickLabelPosition = AxisTickLabelPosition.High;
            yAxis.MajorUnit = 100.0d;
            yAxis.MinorUnit = 50.0d;
            yAxis.DisplayUnit.Unit = AxisBuiltInUnit.Hundreds;
            yAxis.Scaling.Minimum = new AxisBound(100);
            yAxis.Scaling.Maximum = new AxisBound(700);

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

            Assert.AreEqual(AxisTickLabelPosition.High, chart.AxisY.TickLabelPosition);
            Assert.AreEqual(100.0d, chart.AxisY.MajorUnit);
            Assert.AreEqual(50.0d, chart.AxisY.MinorUnit);
            Assert.AreEqual(AxisBuiltInUnit.Hundreds, chart.AxisY.DisplayUnit.Unit);
            Assert.AreEqual(new AxisBound(100), chart.AxisY.Scaling.Minimum);
            Assert.AreEqual(new AxisBound(700), chart.AxisY.Scaling.Maximum);
        }

        [Test]
        public void HideChartAxis()
        {
            //ExStart
            //ExFor:ChartAxis.Hidden
            //ExSummary:Shows how to hide chart axes.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a line chart, clear its demo data, and insert a custom series with decimal values. 
            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);
            Chart chart = shape.Chart;
            chart.Series.Clear();
            chart.Series.Add("AW Series 1",
                new string[] { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" },
                new double[] { 1.2, 0.3, 2.1, 2.9, 4.2 });

            // We can hide chart axes to simplify the appearance of our charts. 
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
            //ExFor:Charts.ChartNumberFormat
            //ExFor:ChartNumberFormat.FormatCode
            //ExFor:Charts.ChartNumberFormat.IsLinkedToSource
            //ExSummary:Shows how to set formatting for chart values.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a column chart, clear its demo data, and insert a custom series with large decimal values. 
            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = shape.Chart;
            chart.Series.Clear();
            chart.Series.Add("Aspose Test Series",
                new string[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
                new double[] { 1900000.0d, 850000.0d, 2100000.0d, 600000.0d, 1500000.0d });

            // Set number format to not group digits with commas. 
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

            // Insert chart
            Shape shape = builder.InsertChart(chartType, 432, 252);
            Chart chart = shape.Chart;

            // Clear demo data
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

            // Insert chart
            Shape shape = builder.InsertChart(ChartType.Surface3D, 432, 252);
            Chart chart = shape.Chart;

            // Clear demo data
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

            // Insert a bubble chart, clear its demo data,
            // and insert a custom series for the X/Y coordinates and thickness of each of the bubbles. 
            Chart chart = builder.InsertChart(ChartType.Bubble, 432, 252).Chart;
            chart.Series.Clear();
            ChartSeries series = chart.Series.Add("Aspose Test Series",
                new[] { 2.9, 3.5, 1.1, 4, 4 },
                new[] { 1.9, 8.5, 2.1, 6, 1.5 },
                new[] { 9, 4.5, 2.5, 8, 5 });

            // Enable data labels and modify their appearance.
            series.HasDataLabels = true;
            ChartDataLabelCollection dataLabels = series.DataLabels;
            dataLabels.ShowBubbleSize = true;
            dataLabels.ShowCategoryName = true;
            dataLabels.ShowSeriesName = true;
            dataLabels.Separator = " & ";

            builder.InsertBreak(BreakType.PageBreak);

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

            // Insert a bubble chart, clear its demo data, and insert a custom series for fractions of a pie chart.
            Chart chart = builder.InsertChart(ChartType.Pie, 432, 252).Chart;
            chart.Series.Clear();
            ChartSeries series = chart.Series.Add("Aspose Test Series",
                new string[] { "Word", "PDF", "Excel" },
                new double[] { 2.7, 3.2, 0.8 });

            // Enable data labels and modify their appearance.
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
        //ExFor:ChartDataLabelCollection.Add(System.Int32)
        //ExFor:ChartDataLabelCollection.Clear
        //ExFor:ChartDataLabelCollection.Count
        //ExFor:ChartDataLabelCollection.GetEnumerator
        //ExFor:ChartDataLabelCollection.Item(System.Int32)
        //ExFor:ChartDataLabelCollection.RemoveAt(System.Int32)
        //ExSummary:Shows how to apply labels to data points in a line chart.
        [Test] //ExSkip
        public void DataLabels()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            // Create a line chart and verify its demo data series.
            Shape chartShape = builder.InsertChart(ChartType.Line, 400, 300);
            Chart chart = chartShape.Chart;

            Assert.AreEqual(3, chart.Series.Count);
            Assert.AreEqual("Series 1", chart.Series[0].Name);

            // Apply data labels to every series in the chart.
            // These labels will display the value of each point of a series in the graph.
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
        /// Apply data labels with custom number format and separator to a number of data points in a series.
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

                // Apply number format and separator
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
        //ExFor:ChartDataPointCollection.Add(System.Int32)
        //ExFor:ChartDataPointCollection.Clear
        //ExFor:ChartDataPointCollection.Count
        //ExFor:ChartDataPointCollection.GetEnumerator
        //ExFor:ChartDataPointCollection.Item(System.Int32)
        //ExFor:ChartDataPointCollection.RemoveAt(System.Int32)
        //ExFor:ChartMarker
        //ExFor:ChartMarker.Size
        //ExFor:ChartMarker.Symbol
        //ExFor:IChartDataPoint
        //ExFor:IChartDataPoint.InvertIfNegative
        //ExFor:IChartDataPoint.Marker
        //ExFor:MarkerSymbol
        //ExSummary:Shows how to work with data points on a line chart.
        [Test]
        public void ChartDataPoint()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a line chart, which will already contain demo data.
            Shape shape = builder.InsertChart(ChartType.Line, 500, 350);
            Chart chart = shape.Chart;

            // Set data points along lines to appear in the form of diamond shapes.
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

            // For a cleaner looking graph, we can remove data points individually.
            chart.Series[1].DataPoints.RemoveAt(2);

            // We can also strip an entire series of its data points at once.
            chart.Series[2].DataPoints.Clear();

            doc.Save(ArtifactsDir + "Charts.ChartDataPoint.docx");
        }

        /// <summary>
        /// Applies a number of data points to a series.
        /// </summary>
        private static void ApplyDataPoints(ChartSeries series, int dataPointsCount, MarkerSymbol markerSymbol, int dataPointSize)
        {
            for (int i = 0; i < dataPointsCount; i++)
            {
                ChartDataPoint point = series.DataPoints.Add(i);
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
            //ExFor:Charts.IChartDataPoint.Explosion
            //ExSummary:Shows how to move the slices of a pie chart away from the center.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a pie chart which will contain demo data.
            Shape shape = builder.InsertChart(ChartType.Pie, 500, 350);
            Chart chart = shape.Chart;

            // "Slices" of a pie chart may be moved away from the center by a distance via the respective data point's Explosion attribute.
            // Add a data point to the first portion of the pie chart and move it away from the center by 10 points. 
            ChartDataPoint dataPoint = chart.Series[0].DataPoints.Add(0);
            dataPoint.Explosion = 10;

            // Displace the second portion by a greater distance.
            dataPoint = chart.Series[0].DataPoints.Add(1);
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
            //ExFor:Charts.ChartDataLabel.ShowBubbleSize
            //ExFor:Charts.IChartDataPoint.Bubble3D
            //ExSummary:Shows how to use 3D effects with bubble charts.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a 3D bubble chart, which will come with demo data by default.
            Shape shape = builder.InsertChart(ChartType.Bubble3D, 500, 350);
            Chart chart = shape.Chart;

            Assert.True(chart.Series[0].Bubble3D);

            // Apply a data label to each bubble that displays its size.
            for (int i = 0; i < 3; i++)
            {
                chart.Series[0].HasDataLabels = true;
                chart.Series[0].DataLabels[i].ShowBubbleSize = true;
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
            Chart chart = AppendChart(builder, ChartType.Column, 300, 300);

            string[] categories = { "Category 1", "Category 2", "Category 3" };

            // Insert two series of decimal values containing a value for each respective category.
            // This column chart will have three groups, each with two columns.
            chart.Series.Add("Series 1", categories, new [] { 76.6, 82.1, 91.6 });
            chart.Series.Add("Series 2", categories, new [] { 64.2, 79.5, 94.0 });

            // Categories are distributed along the X-axis, and values are distributed along the Y-axis.
            Assert.AreEqual(ChartAxisType.Category, chart.AxisX.Type);
            Assert.AreEqual(ChartAxisType.Value, chart.AxisY.Type);

            // 2 -  Area chart with dates distributed along the X-axis:
            chart = AppendChart(builder, ChartType.Area, 300, 300);

            DateTime[] dates = { new DateTime(2014, 3, 31),
                new DateTime(2017, 1, 23),
                new DateTime(2017, 6, 18),
                new DateTime(2019, 11, 22),
                new DateTime(2020, 9, 7)
            };

            // Insert a series with a decimal value for each respective date.
            // The dates will be distributed along a linear X-axis,
            // and the values added to this series will create data points.
            chart.Series.Add("Series 1", dates, new [] { 15.8, 21.5, 22.9, 28.7, 33.1 });

            Assert.AreEqual(ChartAxisType.Category, chart.AxisX.Type);
            Assert.AreEqual(ChartAxisType.Value, chart.AxisY.Type);

            // 3 -  2D scatter plot:
            chart = AppendChart(builder, ChartType.Scatter, 300, 300);

            // Each series will need two decimal arrays of equal length.
            // The first array contains X-values, and the second contains corresponding Y-values
            // of points to be be plotted on the graph.
            chart.Series.Add("Series 1", 
                new[] { 3.1, 3.5, 6.3, 4.1, 2.2, 8.3, 1.2, 3.6 }, 
                new[] { 3.1, 6.3, 4.6, 0.9, 8.5, 4.2, 2.3, 9.9 });
            chart.Series.Add("Series 2", 
                new[] { 2.6, 7.3, 4.5, 6.6, 2.1, 9.3, 0.7, 3.3 }, 
                new[] { 7.1, 6.6, 3.5, 7.8, 7.7, 9.5, 1.3, 4.6 });

            Assert.AreEqual(ChartAxisType.Value, chart.AxisX.Type);
            Assert.AreEqual(ChartAxisType.Value, chart.AxisY.Type);

            // 4 -  Bubble chart:
            chart = AppendChart(builder, ChartType.Bubble, 300, 300);

            // Each series will need three decimal arrays of equal length.
            // The first array contains X-values, the second contains corresponding Y-values,
            // and the third contains diameters of points to be plotted on the graph.
            chart.Series.Add("Series 1", 
                new [] { 1.1, 5.0, 9.8 }, 
                new [] { 1.2, 4.9, 9.9 }, 
                new [] { 2.0, 4.0, 8.0 });

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

            Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
            Chart chart = chartShape.Chart;

            // This column chart currently contains three series of demo data.
            // Each series has four decimal values; one for each of the four categories.
            // This data will be represented by four clusters of three columns.
            ChartSeriesCollection chartData = chart.Series;
            Assert.AreEqual(3, chartData.Count); //ExSkip

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

            // We can also clear all of the chart's data at once with this method.
            // When creating a new chart, this is the way to wipe all the demo data
            // before we can begin working on a clean chart.
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

            // Insert a scatter chart and clear its demo data.
            Shape chartShape = builder.InsertChart(ChartType.Scatter, 450, 300);
            Chart chart = chartShape.Chart;
            chart.Series.Clear();

            // Insert a series with X/Y coordinates for five points.
            chart.Series.Add("Series 1", 
                new[] { 1.0, 2.0, 3.0, 4.0, 5.0 }, 
                new[] { 1.0, 20.0, 400.0, 8000.0, 160000.0 });

            // The scaling of the X-axis is linear by default,
            // displaying evenly incrementing values that cover our X-value range (0, 1, 2, 3...).
            // A linear axis is not ideal for our Y-values,
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

            // Insert a scatter chart, remove default data and populate it with data from a series.
            Shape chartShape = builder.InsertChart(ChartType.Scatter, 450, 300);
            Chart chart = chartShape.Chart;
            chart.Series.Clear();
            chart.Series.Add("Series 1", 
                new[] { 1.1, 5.4, 7.9, 3.5, 2.1, 9.7 }, 
                new[] { 2.1, 0.3, 0.6, 3.3, 1.4, 1.9 });

            // By default, the axis bounds are automatically defined,
            // so that both X and Y-axis ranges are just big enough to encompass every value of every series.
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

            // We can set axis bounds in the form of dates as well, limiting the chart to a time period. 
            // Setting the range to 1980-1990 will cause the two of the series values
            // that are outside of the range to be omitted from the graph.
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

            // Insert a line chart, and move its legend to the top right corner. 
            Shape shape = builder.InsertChart(ChartType.Line, 450, 300);
            ChartLegend legend = shape.Chart.Legend;
            legend.Position = LegendPosition.TopRight;

            // Give other chart elements such as the graph more room by allowing them to overlap the legend.
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

            // Insert a column chart, which contains demo data.
            Shape shape = builder.InsertChart(ChartType.Column, 450, 250);
            Chart chart = shape.Chart;

            // For column charts, the Y-axis crosses at zero by default,
            // which means that columns for all values below zero point down to represent negative values.
            // We can set a different value for the Y-axis crossing. In this case we will set it to 3.
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
            //ExFor:ChartAxis.TickLabelSpacing
            //ExFor:ChartAxis.TickLabelAlignment
            //ExFor:AxisDisplayUnit
            //ExFor:AxisDisplayUnit.CustomUnit
            //ExFor:AxisDisplayUnit.Unit
            //ExSummary:Shows how to manipulate the tick marks and displayed values of a chart axis.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a scatter chart, which contains demo data.
            Shape shape = builder.InsertChart(ChartType.Scatter, 450, 250);
            Chart chart = shape.Chart;

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
            // every minor tick mark at 2.5 units, and configure them to both be inside the graph plot area.
            axis = chart.AxisX;
            axis.MajorTickMark = AxisTickMark.Inside;
            axis.MinorTickMark = AxisTickMark.Inside;
            axis.MajorUnit = 10;
            axis.MinorUnit = 2.5;

            // Set the X-axis bounds so that the X-axis spans 5 major tick marks and 12 minor tick marks.
            axis.Scaling.Minimum = new AxisBound(-10);
            axis.Scaling.Maximum = new AxisBound(30);
            axis.TickLabelAlignment = ParagraphAlignment.Right;

            Assert.AreEqual(1, axis.TickLabelSpacing);
            
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
            Assert.AreEqual(1, axis.TickLabelSpacing);
            Assert.AreEqual(ParagraphAlignment.Right, axis.TickLabelAlignment);
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
    }
}