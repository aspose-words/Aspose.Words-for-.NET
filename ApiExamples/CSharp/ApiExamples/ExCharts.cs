// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
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
            //ExSummary:Shows how to insert a chart and change its title.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a document builder to insert a bar chart
            Shape chartShape = builder.InsertChart(ChartType.Bar, 400, 300);

            // Get the chart object from the containing shape
            Chart chart = chartShape.Chart;
            
            // Set the title text, which appears at the top center of the chart and modify its appearance
            ChartTitle title = chart.Title;
            title.Text = "MyChart";
            title.Overlay = true;
            title.Show = true;

            doc.Save(ArtifactsDir + "Charts.ChartTitle.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.ChartTitle.docx");
            chartShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual(ShapeType.NonPrimitive, chartShape.ShapeType);
            Assert.True(chartShape.HasChart);

            title = chartShape.Chart.Title;

            Assert.AreEqual("MyChart", title.Text);
            Assert.True(title.Overlay);
            Assert.True(title.Show);
        }

        [Test]
        public void DefineNumberFormatForDataLabels()
        {
            //ExStart
            //ExFor:ChartDataLabelCollection.NumberFormat
            //ExFor:ChartNumberFormat.FormatCode
            //ExSummary:Shows how to set number format for the data labels of the entire series.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add chart with default data
            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);
            // Delete default generated series
            shape.Chart.Series.Clear();
            
            ChartSeries series =
                shape.Chart.Series.Add("Aspose Test Series", new[] { "Word", "PDF", "Excel" }, new[] { 2.5, 1.5, 3.5 });

            ChartDataLabelCollection dataLabels = series.DataLabels;
            // Display chart values in the data labels, by default it is false
            dataLabels.ShowValue = true;
            // Set currency format for the data labels of the entire series
            dataLabels.NumberFormat.FormatCode = "\"$\"#,##0.00";

            doc.Save(ArtifactsDir + "Charts.DefineNumberFormatForDataLabels.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.DefineNumberFormatForDataLabels.docx");
            series = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart.Series[0];

            Assert.AreEqual(string.Empty, series.DataLabels.NumberFormat.FormatCode);
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
            //ExSummary:Shows how to insert chart using the axis options for detailed configuration.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert chart.
            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = shape.Chart;

            // Clear demo data.
            chart.Series.Clear();
            chart.Series.Add("Aspose Test Series",
                new[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
                new double[] { 640, 320, 280, 120, 150 });

            // Get chart axes
            ChartAxis xAxis = chart.AxisX;
            ChartAxis yAxis = chart.AxisY;

            // For 2D charts like the one we made, the Z axis is null
            Assert.Null(chart.AxisZ);

            // Set X-axis options
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

            // Set Y-axis options
            yAxis.CategoryType = AxisCategoryType.Automatic;
            yAxis.Crosses = AxisCrosses.Maximum;
            yAxis.ReverseOrder = true;
            yAxis.MajorTickMark = AxisTickMark.Inside;
            yAxis.MinorTickMark = AxisTickMark.Cross;
            yAxis.MajorUnit = 100.0d;
            yAxis.MinorUnit = 20.0d;
            yAxis.TickLabelPosition = AxisTickLabelPosition.NextToAxis;

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

            // Insert chart
            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);
            Chart chart = shape.Chart;
            
            // Clear demo data
            chart.Series.Clear();

            // Fill data
            chart.Series.Add("Aspose Test Series",
                new[]
                {
                    new DateTime(2017, 11, 06), new DateTime(2017, 11, 09), new DateTime(2017, 11, 15),
                    new DateTime(2017, 11, 21), new DateTime(2017, 11, 25), new DateTime(2017, 11, 29)
                },
                new[] { 1.2, 0.3, 2.1, 2.9, 4.2, 5.3 });

            ChartAxis xAxis = chart.AxisX;
            ChartAxis yAxis = chart.AxisY;

            // Set X axis bounds
            xAxis.Scaling.Minimum = new AxisBound(new DateTime(2017, 11, 05).ToOADate());
            xAxis.Scaling.Maximum = new AxisBound(new DateTime(2017, 12, 03));

            // Set major units to a week and minor units to a day
            xAxis.BaseTimeUnit = AxisTimeUnit.Days;
            xAxis.MajorUnit = 7.0d;
            xAxis.MinorUnit = 1.0d;
            xAxis.MajorTickMark = AxisTickMark.Cross;
            xAxis.MinorTickMark = AxisTickMark.Outside;

            // Define Y axis properties
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

            // Insert chart
            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);
            Chart chart = shape.Chart;

            // Hide both the X and Y axes
            chart.AxisX.Hidden = true;
            chart.AxisY.Hidden = true;

            // Clear demo data
            chart.Series.Clear();
            chart.Series.Add("AW Series 1",
                new string[] { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" },
                new double[] { 1.2, 0.3, 2.1, 2.9, 4.2 });

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

            // Insert chart
            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = shape.Chart;

            // Clear demo data and replace it with a new custom chart series
            chart.Series.Clear();
            chart.Series.Add("Aspose Test Series",
                new string[] { "Word", "PDF", "Excel", "GoogleDocs", "Note" },
                new double[] { 1900000.0d, 850000.0d, 2100000.0d, 600000.0d, 1500000.0d });

            // Set number format
            chart.AxisY.NumberFormat.FormatCode = "#,##0";

            // Set this to override the above value and draw the number format from the source cell
            Assert.False(chart.AxisY.NumberFormat.IsLinkedToSource);

            doc.Save(ArtifactsDir + "Charts.SetNumberFormatToChartAxis.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.SetNumberFormatToChartAxis.docx");
            chart = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart;

            Assert.AreEqual("#,##0", chart.AxisY.NumberFormat.FormatCode);
        }

        // Note: Tests below used for verification conversion docx to pdf and the correct display.
        // For now, the results check manually.
        [Test]
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
        public void ChartDataLabelCollection()
        {
            //ExStart
            //ExFor:ChartDataLabelCollection.ShowBubbleSize
            //ExFor:ChartDataLabelCollection.ShowCategoryName
            //ExFor:ChartDataLabelCollection.ShowSeriesName
            //ExFor:ChartDataLabelCollection.Separator
            //ExFor:ChartDataLabelCollection.ShowLeaderLines
            //ExFor:ChartDataLabelCollection.ShowLegendKey
            //ExFor:ChartDataLabelCollection.ShowPercentage
            //ExFor:ChartDataLabelCollection.ShowValue
            //ExSummary:Shows how to set default values for the data labels.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert bubble chart
            Chart chart = builder.InsertChart(ChartType.Bubble, 432, 252).Chart;
            // Clear demo data
            chart.Series.Clear();
            
            ChartSeries bubbleChartSeries = chart.Series.Add("Aspose Test Series",
                new[] { 2.9, 3.5, 1.1, 4, 4 },
                new[] { 1.9, 8.5, 2.1, 6, 1.5 },
                new[] { 9, 4.5, 2.5, 8, 5 });

            // Set default values for the bubble chart data labels
            ChartDataLabelCollection bubbleChartDataLabels = bubbleChartSeries.DataLabels;
            bubbleChartDataLabels.ShowBubbleSize = true;
            bubbleChartDataLabels.ShowCategoryName = true;
            bubbleChartDataLabels.ShowSeriesName = true;
            bubbleChartDataLabels.Separator = " - ";

            builder.InsertBreak(BreakType.PageBreak);

            // Insert pie chart
            Shape shapeWithPieChart = builder.InsertChart(ChartType.Pie, 432, 252);
            // Clear demo data
            shapeWithPieChart.Chart.Series.Clear();

            ChartSeries pieChartSeries = shapeWithPieChart.Chart.Series.Add("Aspose Test Series",
                new string[] { "Word", "PDF", "Excel" },
                new double[] { 2.7, 3.2, 0.8 });

            // Set default values for the pie chart data labels
            ChartDataLabelCollection pieChartDataLabels = pieChartSeries.DataLabels;
            pieChartDataLabels.ShowLeaderLines = true;
            pieChartDataLabels.ShowLegendKey = true;
            pieChartDataLabels.ShowPercentage = true;
            pieChartDataLabels.ShowValue = true;

            doc.Save(ArtifactsDir + "Charts.ChartDataLabelCollection.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.ChartDataLabelCollection.docx");
            bubbleChartDataLabels = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart.Series[0].DataLabels;

            Assert.False(bubbleChartDataLabels.ShowBubbleSize);
            Assert.False(bubbleChartDataLabels.ShowCategoryName);
            Assert.False(bubbleChartDataLabels.ShowSeriesName);
            Assert.AreEqual(",", bubbleChartDataLabels.Separator);

            pieChartDataLabels = ((Shape)doc.GetChild(NodeType.Shape, 1, true)).Chart.Series[0].DataLabels;

            Assert.False(pieChartDataLabels.ShowLeaderLines);
            Assert.False(pieChartDataLabels.ShowLegendKey);
            Assert.False(pieChartDataLabels.ShowPercentage);
            Assert.False(pieChartDataLabels.ShowValue);
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
        //ExFor:ChartDataLabelCollection
        //ExFor:ChartDataLabelCollection.Add(System.Int32)
        //ExFor:ChartDataLabelCollection.Clear
        //ExFor:ChartDataLabelCollection.Count
        //ExFor:ChartDataLabelCollection.GetEnumerator
        //ExFor:ChartDataLabelCollection.Item(System.Int32)
        //ExFor:ChartDataLabelCollection.RemoveAt(System.Int32)
        //ExSummary:Shows how to apply labels to data points in a chart.
        [Test] //ExSkip
        public void ChartDataLabels()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            // Use a document builder to insert a bar chart
            Shape chartShape = builder.InsertChart(ChartType.Line, 400, 300);

            // Get the chart object from the containing shape
            Chart chart = chartShape.Chart;

            // The chart already contains demo data comprised of 3 series each with 4 categories
            Assert.AreEqual(3, chart.Series.Count);
            Assert.AreEqual("Series 1", chart.Series[0].Name);

            // Apply data labels to every series in the graph
            foreach (ChartSeries series in chart.Series)
            {
                ApplyDataLabels(series, 4, "000.0", ", ");
                Assert.AreEqual(4, series.DataLabels.Count);
            }

            // Get the enumerator for a data label collection
            using (IEnumerator<ChartDataLabel> enumerator = chart.Series[0].DataLabels.GetEnumerator())
            {
                // And use it to go over all the data labels in one series and change their separator
                while (enumerator.MoveNext())
                {
                    Assert.AreEqual(", ", enumerator.Current.Separator);
                    enumerator.Current.Separator = " & ";
                }
            }

            // If the chart looks too busy, we can remove data labels one by one
            chart.Series[1].DataLabels[2].ClearFormat();

            // We can also clear an entire data label collection for one whole series
            chart.Series[2].DataLabels.ClearFormat();

            doc.Save(ArtifactsDir + "Charts.ChartDataLabels.docx");
        }

        /// <summary>
        /// Apply uniform data labels with custom number format and separator to a number (determined by labelsCount) of data points in a series
        /// </summary>
        private static void ApplyDataLabels(ChartSeries series, int labelsCount, string numberFormat, string separator)
        {
            for (int i = 0; i < labelsCount; i++)
            {
                series.HasDataLabels = true;
                Assert.False(series.DataLabels[i].IsVisible);

                // Edit the appearance of the new data label
                series.DataLabels[i].ShowCategoryName = true;
                series.DataLabels[i].ShowSeriesName = true;
                series.DataLabels[i].ShowValue = true;
                series.DataLabels[i].ShowLeaderLines = true;
                series.DataLabels[i].ShowLegendKey = true;
                series.DataLabels[i].ShowPercentage = false;
                Assert.False(series.DataLabels[i].ShowDataLabelsRange);

                // Apply number format and separator
                series.DataLabels[i].NumberFormat.FormatCode = numberFormat;
                series.DataLabels[i].Separator = separator;

                // The label automatically becomes visible
                Assert.True(series.DataLabels[i].IsVisible);
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
        //ExSummary:Shows how to customize chart data points.
        [Test]
        public void ChartDataPoint()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a line chart, which will have default data that we will use
            Shape shape = builder.InsertChart(ChartType.Line, 500, 350);
            Chart chart = shape.Chart;

            // Apply diamond-shaped data points to the line of the first series
            foreach (ChartSeries series in chart.Series) 
                ApplyDataPoints(series, 4, MarkerSymbol.Diamond, 15);

            // We can further decorate a series line by smoothing it
            chart.Series[0].Smooth = true;

            // Get the enumerator for the data point collection from one series
            using (IEnumerator<ChartDataPoint> enumerator = chart.Series[0].DataPoints.GetEnumerator())
            {
                // And use it to go over all the data labels in one series and change their separator
                while (enumerator.MoveNext())
                {
                    Assert.False(enumerator.Current.InvertIfNegative);
                }
            }

            // If the chart looks too busy, we can remove data points one by one
            chart.Series[1].DataPoints.RemoveAt(2);

            // We can also clear an entire data point collection for one whole series
            chart.Series[2].DataPoints.Clear();

            doc.Save(ArtifactsDir + "Charts.ChartDataPoint.docx");
        }

        /// <summary>
        /// Applies a number of data points to a series
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
            //ExSummary:Shows how to manipulate the position of the portions of a pie chart.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Pie, 500, 350);
            Chart chart = shape.Chart;

            // In a pie chart, the portions are the data points, which cannot have markers or sizes applied to them
            // However, we can set this variable to move any individual "slice" away from the center of the chart
            ChartDataPoint cdp = chart.Series[0].DataPoints.Add(0);
            cdp.Explosion = 10;

            cdp = chart.Series[0].DataPoints.Add(1);
            cdp.Explosion = 40;

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

            // Insert a bubble chart with 3D effects on each bubble
            Shape shape = builder.InsertChart(ChartType.Bubble3D, 500, 350);
            Chart chart = shape.Chart;

            Assert.True(chart.Series[0].Bubble3D);

            // Apply a data label to each bubble that displays the size of its bubble
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
        //ExSummary:Shows how to pick an appropriate graph type for a chart series.
        [Test] //ExSkip
        public void ChartSeriesCollection()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            // There are 4 ways of populating a chart's series collection
            // 1: Each series has a string array of categories, each with a corresponding data value
            // Some of the other possible applications are bar, column, line and surface charts
            Chart chart = AppendChart(builder, ChartType.Column, 300, 300);

            // Create and name 3 categories with a string array
            string[] categories = { "Category 1", "Category 2", "Category 3" };

            // Create 2 series of data, each with one point for every category
            // This will generate a column graph with 3 clusters of 2 bars
            chart.Series.Add("Series 1", categories, new [] { 76.6, 82.1, 91.6 });
            chart.Series.Add("Series 2", categories, new [] { 64.2, 79.5, 94.0 });

            // Categories are distributed along the X-axis while values are distributed along the Y-axis
            Assert.AreEqual(ChartAxisType.Category, chart.AxisX.Type);
            Assert.AreEqual(ChartAxisType.Value, chart.AxisY.Type);

            // 2: Each series will have a collection of dates with a corresponding value for each date
            // Area, radar and stock charts are some of the appropriate chart types for this
            chart = AppendChart(builder, ChartType.Area, 300, 300);

            // Create a collection of dates to serve as categories
            DateTime[] dates = { new DateTime(2014, 3, 31),
                new DateTime(2017, 1, 23),
                new DateTime(2017, 6, 18),
                new DateTime(2019, 11, 22),
                new DateTime(2020, 9, 7)
            };

            // Add one series with one point for each date
            // Our sporadic dates will be distributed along the X-axis in a linear fashion 
            chart.Series.Add("Series 1", dates, new [] { 15.8, 21.5, 22.9, 28.7, 33.1 });

            // 3: Each series will take two data arrays
            // Appropriate for scatter plots
            chart = AppendChart(builder, ChartType.Scatter, 300, 300);

            // In each series, the first array contains the X-coordinates and the second contains respective Y-coordinates of points
            chart.Series.Add("Series 1", new[] { 3.1, 3.5, 6.3, 4.1, 2.2, 8.3, 1.2, 3.6 }, new[] { 3.1, 6.3, 4.6, 0.9, 8.5, 4.2, 2.3, 9.9 });
            chart.Series.Add("Series 2", new[] { 2.6, 7.3, 4.5, 6.6, 2.1, 9.3, 0.7, 3.3 }, new[] { 7.1, 6.6, 3.5, 7.8, 7.7, 9.5, 1.3, 4.6 });

            // Both axes are value axes in this case
            Assert.AreEqual(ChartAxisType.Value, chart.AxisX.Type);
            Assert.AreEqual(ChartAxisType.Value, chart.AxisY.Type);

            // 4: Each series will be built from three data arrays, used for bubble charts
            chart = AppendChart(builder, ChartType.Bubble, 300, 300);

            // The first two arrays contain X/Y coordinates like above and the third determines the thickness of each point
            chart.Series.Add("Series 1", new [] { 1.1, 5.0, 9.8 }, new [] { 1.2, 4.9, 9.9 }, new [] { 2.0, 4.0, 8.0 });

            doc.Save(ArtifactsDir + "Charts.ChartSeriesCollection.docx");
        }
        
        /// <summary>
        /// Get the DocumentBuilder to insert a chart of a specified ChartType, width and height and clean out its default data
        /// </summary>
        private static Chart AppendChart(DocumentBuilder builder, ChartType chartType, double width, double height)
        {
            Shape chartShape = builder.InsertChart(chartType, width, height);
            Chart chart = chartShape.Chart;
            chart.Series.Clear();

            Assert.AreEqual(0, chart.Series.Count);

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
            //ExSummary:Shows how to work with a chart's data collection.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a document builder to insert a bar chart
            Shape chartShape = builder.InsertChart(ChartType.Column, 400, 300);
            Chart chart = chartShape.Chart;

            // All charts come with demo data
            // This column chart currently has 3 series with 4 categories, which means 4 clusters, 3 columns in each
            ChartSeriesCollection chartData = chart.Series;
            Assert.AreEqual(3, chartData.Count); //ExSkip

            // Iterate through the series with an enumerator and print their names
            using (IEnumerator<ChartSeries> enumerator = chart.Series.GetEnumerator())
            {
                // And use it to go over all the data labels in one series and change their separator
                while (enumerator.MoveNext())
                {
                    Console.WriteLine(enumerator.Current.Name);
                }
            }

            // We can add new data by adding a new series to the collection, with categories and data
            // We will match the existing category/series names in the demo data and add a 4th column to each column cluster
            string[] categories = { "Category 1", "Category 2", "Category 3", "Category 4" };
            chart.Series.Add("Series 4", categories, new[] { 4.4, 7.0, 3.5, 2.1 });
            Assert.AreEqual(4, chartData.Count); //ExSkip
            Assert.AreEqual("Series 4", chartData[3].Name); //ExSkip

            // We can remove series by index
            chartData.RemoveAt(2);
            Assert.AreEqual(3, chartData.Count); //ExSkip
            Assert.AreEqual("Series 4", chartData[2].Name); //ExSkip

            // We can also remove out all the series
            // This leaves us with an empty graph and is a convenient way of wiping out demo data
            chartData.Clear();
            Assert.AreEqual(0, chartData.Count); //ExSkip
        }

        [Test]
        public void AxisScaling()
        {
            //ExStart
            //ExFor:AxisScaleType
            //ExFor:AxisScaling
            //ExFor:AxisScaling.LogBase
            //ExFor:AxisScaling.Type
            //ExSummary:Shows how to set up logarithmic axis scaling.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a scatter chart and clear its default data series
            Shape chartShape = builder.InsertChart(ChartType.Scatter, 450, 300);
            Chart chart = chartShape.Chart;
            chart.Series.Clear();

            // Insert a series with X/Y coordinates for 5 points
            chart.Series.Add("Series 1", new[] { 1.0, 2.0, 3.0, 4.0, 5.0 }, new[] { 1.0, 20.0, 400.0, 8000.0, 160000.0 });

            // The scaling of the X axis is linear by default, which means it will display "0, 1, 2, 3..."
            // Linear axis scaling is suitable for our X-values, but our Y-values call for a logarithmic scale to be represented accurately on a graph 
            // We can set the scaling of the Y-axis to Logarithmic with a base of 20
            // The Y-axis will now display "1, 20, 400, 8000...", which is ideal for accurate representation of this set of Y-values
            chart.AxisY.Scaling.Type = AxisScaleType.Logarithmic;
            chart.AxisY.Scaling.LogBase = 20.0d;

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

            // Insert a scatter chart, remove default data and populate it with data from a ChartSeries
            Shape chartShape = builder.InsertChart(ChartType.Scatter, 450, 300);
            Chart chart = chartShape.Chart;
            chart.Series.Clear();
            chart.Series.Add("Series 1", new[] { 1.1, 5.4, 7.9, 3.5, 2.1, 9.7 }, new[] { 2.1, 0.3, 0.6, 3.3, 1.4, 1.9 });

            // By default, the axis bounds are automatically defined so all the series data within the table is included
            Assert.True(chart.AxisX.Scaling.Minimum.IsAuto);

            // If we wish to set our own scale bounds, we need to replace them with new ones
            // Both the axis rulers will go from 0 to 10
            chart.AxisX.Scaling.Minimum = new AxisBound(0);
            chart.AxisX.Scaling.Maximum = new AxisBound(10);
            chart.AxisY.Scaling.Minimum = new AxisBound(0);
            chart.AxisY.Scaling.Maximum = new AxisBound(10);

            // These are custom and not defined automatically
            Assert.False(chart.AxisX.Scaling.Minimum.IsAuto);
            Assert.False(chart.AxisY.Scaling.Minimum.IsAuto);

            // Create a line graph
            chartShape = builder.InsertChart(ChartType.Line, 450, 300);
            chart = chartShape.Chart;
            chart.Series.Clear();

            // Create a collection of dates, which will make up the X axis
            DateTime[] dates = { new DateTime(1973, 5, 11),
                new DateTime(1981, 2, 4),
                new DateTime(1985, 9, 23),
                new DateTime(1989, 6, 28),
                new DateTime(1994, 12, 15)
            };

            // Assign a Y-value for each date 
            chart.Series.Add("Series 1", dates, new[] { 3.0, 4.7, 5.9, 7.1, 8.9 });

            // These particular bounds will cut off categories from before 1980 and from 1990 and onwards
            // This narrows the amount of categories and values in the viewport from 5 to 3
            // Note that the graph still contains the out-of-range data because we can see the line tend towards it
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

            // Insert a line graph
            Shape shape = builder.InsertChart(ChartType.Line, 450, 300);

            // Get its legend
            ChartLegend legend = shape.Chart.Legend;

            // By default, other elements of a chart will not overlap with its legend
            Assert.False(legend.Overlay);

            // We can move its position by setting this attribute
            legend.Position = LegendPosition.TopRight;

            doc.Save(ArtifactsDir + "Charts.ChartLegend.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.ChartLegend.docx");

            legend = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart.Legend;

            Assert.False(legend.Overlay);
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

            // Insert a column chart, which is populated by default values
            Shape shape = builder.InsertChart(ChartType.Column, 450, 250);
            Chart chart = shape.Chart;

            // Get the Y-axis to cross at a value of 3.0, making 3.0 the new Y-zero of our column chart
            // This effectively means that all the columns with Y-values about 3.0 will be above the Y-centre and point up,
            // while ones below 3.0 will point down
            ChartAxis axis = chart.AxisX;
            axis.AxisBetweenCategories = true;
            axis.Crosses = AxisCrosses.Custom;
            axis.CrossesAt = 3.0;

            doc.Save(ArtifactsDir + "Charts.AxisCross.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.AxisCross.docx");
            axis = ((Shape)doc.GetChild(NodeType.Shape, 0, true)).Chart.AxisX;

            Assert.True(axis.AxisBetweenCategories);
            Assert.AreEqual(AxisCrosses.Custom, axis.Crosses);
            Assert.AreEqual(3.0, axis.CrossesAt);
        }

        [Test]
        public void ChartAxisDisplayUnit()
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

            // Insert a scatter chart, which is populated by default values
            Shape shape = builder.InsertChart(ChartType.Scatter, 450, 250);
            Chart chart = shape.Chart;

            // Set they Y axis to show major ticks every at every 10 units and minor ticks at every 1 units
            ChartAxis axis = chart.AxisY;
            axis.MajorTickMark = AxisTickMark.Outside;
            axis.MinorTickMark = AxisTickMark.Outside;

            axis.MajorUnit = 10.0d;
            axis.MinorUnit = 1.0d;

            // Stretch out the bounds of the axis out to show 3 major ticks and 27 minor ticks
            axis.Scaling.Minimum = new AxisBound(-10);
            axis.Scaling.Maximum = new AxisBound(20);

            // Do the same for the X-axis
            axis = chart.AxisX;
            axis.MajorTickMark = AxisTickMark.Inside;
            axis.MinorTickMark = AxisTickMark.Inside;
            axis.MajorUnit = 10.0;
            axis.Scaling.Minimum = new AxisBound(-10);
            axis.Scaling.Maximum = new AxisBound(30);

            // We can also use this attribute to set minor tick spacing
            axis.TickLabelSpacing = 2;
            // We can define text alignment when axis tick labels are multi-line
            // MS Word aligns them to the center by default
            axis.TickLabelAlignment = ParagraphAlignment.Right;

            // Get the axis to display values, but in millions
            axis.DisplayUnit.Unit = AxisBuiltInUnit.Millions;
            Assert.AreEqual(AxisBuiltInUnit.Millions, axis.DisplayUnit.Unit); //ExSkip

            // Besides the built-in axis units we can choose from,
            // we can also set the axis to display values in some custom denomination, using the following attribute
            // The statement below is equivalent to the one above
            axis.DisplayUnit.CustomUnit = 1000000.0;
            Assert.AreEqual(AxisBuiltInUnit.Custom, axis.DisplayUnit.Unit); //ExSkip

            doc.Save(ArtifactsDir + "Charts.ChartAxisDisplayUnit.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "Charts.ChartAxisDisplayUnit.docx");
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

            Assert.AreEqual(AxisTickMark.Outside, axis.MajorTickMark);
            Assert.AreEqual(AxisTickMark.Outside, axis.MinorTickMark);
            Assert.AreEqual(10.0d, axis.MajorUnit);
            Assert.AreEqual(1.0d, axis.MinorUnit);
            Assert.AreEqual(-10.0d, axis.Scaling.Minimum.Value);
            Assert.AreEqual(20.0d, axis.Scaling.Maximum.Value);
        }
    }
}