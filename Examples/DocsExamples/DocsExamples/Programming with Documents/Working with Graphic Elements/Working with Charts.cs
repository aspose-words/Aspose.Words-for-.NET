using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Working_with_Graphic_Elements
{
    internal class WorkingWithCharts : DocsExamplesBase
    {
        [Test]
        public void FormatNumberOfDataLabel()
        {
            //ExStart:FormatNumberOfDataLabel
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);

            Chart chart = shape.Chart;
            chart.Title.Text = "Data Labels With Different Number Format";

            // Delete default generated series.
            chart.Series.Clear();

            ChartSeries series1 = chart.Series.Add("Aspose Series 1", 
                new string[] { "Category 1", "Category 2", "Category 3" }, 
                new double[] { 2.5, 1.5, 3.5 });
            
            series1.HasDataLabels = true;
            series1.DataLabels.ShowValue = true;
            series1.DataLabels[0].NumberFormat.FormatCode = "\"$\"#,##0.00";
            series1.DataLabels[1].NumberFormat.FormatCode = "dd/mm/yyyy";
            series1.DataLabels[2].NumberFormat.FormatCode = "0.00%";

            // Or you can set format code to be linked to a source cell,
            // in this case NumberFormat will be reset to general and inherited from a source cell.
            series1.DataLabels[2].NumberFormat.IsLinkedToSource = true;

            doc.Save(ArtifactsDir + "WorkingWithCharts.FormatNumberOfDataLabel.docx");
            //ExEnd:FormatNumberOfDataLabel
        }

        [Test]
        public void CreateChartUsingShape()
        {
            //ExStart:CreateChartUsingShape
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);

            Chart chart = shape.Chart;
            chart.Title.Show = true;
            chart.Title.Text = "Line Chart Title";
            chart.Title.Overlay = false;

            // Please note if null or empty value is specified as title text, auto generated title will be shown.

            chart.Legend.Position = LegendPosition.Left;
            chart.Legend.Overlay = true;
            
            doc.Save(ArtifactsDir + "WorkingWithCharts.CreateChartUsingShape.docx");
            //ExEnd:CreateChartUsingShape
        }

        [Test]
        public void InsertSimpleColumnChart()
        {
            //ExStart:InsertSimpleColumnChart
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // You can specify different chart types and sizes.
            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);

            Chart chart = shape.Chart;
            //ExStart:ChartSeriesCollection 
            ChartSeriesCollection seriesColl = chart.Series;

            Console.WriteLine(seriesColl.Count);
            //ExEnd:ChartSeriesCollection 

            // Delete default generated series.
            seriesColl.Clear();

            // Create category names array, in this example we have two categories.
            string[] categories = new string[] { "Category 1", "Category 2" };

            // Please note, data arrays must not be empty and arrays must be the same size.
            seriesColl.Add("Aspose Series 1", categories, new double[] { 1, 2 });
            seriesColl.Add("Aspose Series 2", categories, new double[] { 3, 4 });
            seriesColl.Add("Aspose Series 3", categories, new double[] { 5, 6 });
            seriesColl.Add("Aspose Series 4", categories, new double[] { 7, 8 });
            seriesColl.Add("Aspose Series 5", categories, new double[] { 9, 10 });

            doc.Save(ArtifactsDir + "WorkingWithCharts.InsertSimpleColumnChart.docx");
            //ExEnd:InsertSimpleColumnChart
        }

        [Test]
        public void InsertColumnChart()
        {
            //ExStart:InsertColumnChart
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);

            Chart chart = shape.Chart;
            chart.Series.Add("Aspose Series 1", new string[] { "Category 1", "Category 2" }, new double[] { 1, 2 });

            doc.Save(ArtifactsDir + "WorkingWithCharts.InsertColumnChart.docx");
            //ExEnd:InsertColumnChart
        }

        [Test]
        public void InsertAreaChart()
        {
            //ExStart:InsertAreaChart
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Area, 432, 252);

            Chart chart = shape.Chart;
            chart.Series.Add("Aspose Series 1", new []
                {
                    new DateTime(2002, 05, 01),
                    new DateTime(2002, 06, 01),
                    new DateTime(2002, 07, 01),
                    new DateTime(2002, 08, 01),
                    new DateTime(2002, 09, 01)
                }, 
                new double[] { 32, 32, 28, 12, 15 });
            
            doc.Save(ArtifactsDir + "WorkingWithCharts.InsertAreaChart.docx");
            //ExEnd:InsertAreaChart
        }

        [Test]
        public void InsertBubbleChart()
        {
            //ExStart:InsertBubbleChart
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Bubble, 432, 252);

            Chart chart = shape.Chart;
            chart.Series.Add("Aspose Series 1", new double[] { 0.7, 1.8, 2.6 }, new double[] { 2.7, 3.2, 0.8 },
                new double[] { 10, 4, 8 });
            
            doc.Save(ArtifactsDir + "WorkingWithCharts.InsertBubbleChart.docx");
            //ExEnd:InsertBubbleChart
        }

        [Test]
        public void InsertScatterChart()
        {
            //ExStart:InsertScatterChart
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Scatter, 432, 252);

            Chart chart = shape.Chart;
            chart.Series.Add("Aspose Series 1", new double[] { 0.7, 1.8, 2.6 }, new double[] { 2.7, 3.2, 0.8 });

            doc.Save(ArtifactsDir + "WorkingWithCharts.InsertScatterChart.docx");
            //ExEnd:InsertScatterChart
        }

        [Test]
        public void DefineXYAxisProperties()
        {
            //ExStart:DefineXYAxisProperties
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert chart
            Shape shape = builder.InsertChart(ChartType.Area, 432, 252);

            Chart chart = shape.Chart;

            chart.Series.Clear();

            chart.Series.Add("Aspose Series 1",
                new DateTime[]
                {
                    new DateTime(2002, 01, 01), new DateTime(2002, 06, 01), new DateTime(2002, 07, 01),
                    new DateTime(2002, 08, 01), new DateTime(2002, 09, 01)
                },
                new double[] { 640, 320, 280, 120, 150 });

            ChartAxis xAxis = chart.AxisX;
            ChartAxis yAxis = chart.AxisY;

            // Change the X axis to be category instead of date, so all the points will be put with equal interval on the X axis.
            xAxis.CategoryType = AxisCategoryType.Category;
            xAxis.Crosses = AxisCrosses.Custom;
            xAxis.CrossesAt = 3; // Measured in display units of the Y axis (hundreds).
            xAxis.ReverseOrder = true;
            xAxis.MajorTickMark = AxisTickMark.Cross;
            xAxis.MinorTickMark = AxisTickMark.Outside;
            xAxis.TickLabelOffset = 200;

            yAxis.TickLabelPosition = AxisTickLabelPosition.High;
            yAxis.MajorUnit = 100;
            yAxis.MinorUnit = 50;
            yAxis.DisplayUnit.Unit = AxisBuiltInUnit.Hundreds;
            yAxis.Scaling.Minimum = new AxisBound(100);
            yAxis.Scaling.Maximum = new AxisBound(700);

            doc.Save(ArtifactsDir + "WorkingWithCharts.DefineXYAxisProperties.docx");
            //ExEnd:DefineXYAxisProperties
        }

        [Test]
        public void DateTimeValuesToAxis()
        {
            //ExStart:SetDateTimeValuesToAxis
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = shape.Chart;

            chart.Series.Clear();

            chart.Series.Add("Aspose Series 1",
                new[]
                {
                    new DateTime(2017, 11, 06), new DateTime(2017, 11, 09), new DateTime(2017, 11, 15),
                    new DateTime(2017, 11, 21), new DateTime(2017, 11, 25), new DateTime(2017, 11, 29)
                },
                new double[] { 1.2, 0.3, 2.1, 2.9, 4.2, 5.3 });

            ChartAxis xAxis = chart.AxisX;
            xAxis.Scaling.Minimum = new AxisBound(new DateTime(2017, 11, 05).ToOADate());
            xAxis.Scaling.Maximum = new AxisBound(new DateTime(2017, 12, 03).ToOADate());

            // Set major units to a week and minor units to a day.
            xAxis.MajorUnit = 7;
            xAxis.MinorUnit = 1;
            xAxis.MajorTickMark = AxisTickMark.Cross;
            xAxis.MinorTickMark = AxisTickMark.Outside;

            doc.Save(ArtifactsDir + "WorkingWithCharts.DateTimeValuesToAxis.docx");
            //ExEnd:SetDateTimeValuesToAxis
        }

        [Test]
        public void NumberFormatForAxis()
        {
            //ExStart:SetNumberFormatForAxis
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);

            Chart chart = shape.Chart;

            chart.Series.Clear();

            chart.Series.Add("Aspose Series 1",
                new string[] { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" },
                new double[] { 1900000, 850000, 2100000, 600000, 1500000 });

            chart.AxisY.NumberFormat.FormatCode = "#,##0";

            doc.Save(ArtifactsDir + "WorkingWithCharts.NumberFormatForAxis.docx");
            //ExEnd:SetNumberFormatForAxis
        }

        [Test]
        public void BoundsOfAxis()
        {
            //ExStart:SetboundsOfAxis
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);

            Chart chart = shape.Chart;

            chart.Series.Clear();

            chart.Series.Add("Aspose Series 1",
                new string[] { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" },
                new double[] { 1.2, 0.3, 2.1, 2.9, 4.2 });

            chart.AxisY.Scaling.Minimum = new AxisBound(0);
            chart.AxisY.Scaling.Maximum = new AxisBound(6);

            doc.Save(ArtifactsDir + "WorkingWithCharts.BoundsOfAxis.docx");
            //ExEnd:SetboundsOfAxis
        }

        [Test]
        public void IntervalUnitBetweenLabelsOnAxis()
        {
            //ExStart:SetIntervalUnitBetweenLabelsOnAxis
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);

            Chart chart = shape.Chart;

            chart.Series.Clear();

            chart.Series.Add("Aspose Series 1",
                new string[] { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" },
                new double[] { 1.2, 0.3, 2.1, 2.9, 4.2 });

            chart.AxisX.TickLabelSpacing = 2;

            doc.Save(ArtifactsDir + "WorkingWithCharts.IntervalUnitBetweenLabelsOnAxis.docx");
            //ExEnd:SetIntervalUnitBetweenLabelsOnAxis
        }

        [Test]
        public void HideChartAxis()
        {
            //ExStart:HideChartAxis
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);

            Chart chart = shape.Chart;
            
            chart.Series.Clear();
            
            chart.Series.Add("Aspose Series 1",
                new string[] { "Item 1", "Item 2", "Item 3", "Item 4", "Item 5" },
                new double[] { 1.2, 0.3, 2.1, 2.9, 4.2 });
            
            chart.AxisY.Hidden = true;

            doc.Save(ArtifactsDir + "WorkingWithCharts.HideChartAxis.docx");
            //ExEnd:HideChartAxis
        }

        [Test]
        public void TickMultiLineLabelAlignment()
        {
            //ExStart:TickMultiLineLabelAlignment
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Scatter, 450, 250);

            ChartAxis axis = shape.Chart.AxisX;
            // This property has effect only for multi-line labels.
            axis.TickLabelAlignment = ParagraphAlignment.Right;

            doc.Save(ArtifactsDir + "WorkingWithCharts.TickMultiLineLabelAlignment.docx");
            //ExEnd:TickMultiLineLabelAlignment
        }

        [Test]
        public void ChartDataLabel()
        {
            //ExStart:WorkWithChartDataLabel
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Bar, 432, 252);

            Chart chart = shape.Chart;
            ChartSeries series0 = shape.Chart.Series[0];

            ChartDataLabelCollection labels = series0.DataLabels;
            labels.ShowLegendKey = true;
            // By default, when you add data labels to the data points in a pie chart, leader lines are displayed for data labels that are
            // positioned far outside the end of data points. Leader lines create a visual connection between a data label and its 
            // corresponding data point.
            labels.ShowLeaderLines = true;
            labels.ShowCategoryName = false;
            labels.ShowPercentage = false;
            labels.ShowSeriesName = true;
            labels.ShowValue = true;
            labels.Separator = "/";
            labels.ShowValue = true;
            
            doc.Save(ArtifactsDir + "WorkingWithCharts.ChartDataLabel.docx");
            //ExEnd:WorkWithChartDataLabel
        }

        [Test]
        public void DefaultOptionsForDataLabels()
        {
            //ExStart:DefaultOptionsForDataLabels
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Pie, 432, 252);

            Chart chart = shape.Chart;

            chart.Series.Clear();

            ChartSeries series = chart.Series.Add("Aspose Series 1",
                new string[] { "Category 1", "Category 2", "Category 3" },
                new double[] { 2.7, 3.2, 0.8 });

            ChartDataLabelCollection labels = series.DataLabels;
            labels.ShowPercentage = true;
            labels.ShowValue = true;
            labels.ShowLeaderLines = false;
            labels.Separator = " - ";

            doc.Save(ArtifactsDir + "WorkingWithCharts.DefaultOptionsForDataLabels.docx");
            //ExEnd:DefaultOptionsForDataLabels
        }

        [Test]
        public void SingleChartDataPoint()
        {
            //ExStart:WorkWithSingleChartDataPoint
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);

            Chart chart = shape.Chart;
            ChartSeries series0 = chart.Series[0];
            ChartSeries series1 = chart.Series[1];

            ChartDataPointCollection dataPointCollection = series0.DataPoints;
            ChartDataPoint dataPoint00 = dataPointCollection.Add(0);
            ChartDataPoint dataPoint01 = dataPointCollection.Add(1);

            dataPoint00.Explosion = 50;
            dataPoint00.Marker.Symbol = MarkerSymbol.Circle;
            dataPoint00.Marker.Size = 15;

            dataPoint01.Marker.Symbol = MarkerSymbol.Diamond;
            dataPoint01.Marker.Size = 20;

            ChartDataPoint dataPoint12 = series1.DataPoints.Add(2);
            dataPoint12.InvertIfNegative = true;
            dataPoint12.Marker.Symbol = MarkerSymbol.Star;
            dataPoint12.Marker.Size = 20;

            doc.Save(ArtifactsDir + "WorkingWithCharts.SingleChartDataPoint.docx");
            //ExEnd:WorkWithSingleChartDataPoint
        }

        [Test]
        public void SingleChartSeries()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);

            Chart chart = shape.Chart;

            //ExStart:WorkWithSingleChartSeries
            ChartSeries series0 = chart.Series[0];
            ChartSeries series1 = chart.Series[1];

            series0.Name = "Chart Series Name 1";
            series1.Name = "Chart Series Name 2";

            // You can also specify whether the line connecting the points on the chart shall be smoothed using Catmull-Rom splines.
            series0.Smooth = true;
            series1.Smooth = true;
            //ExEnd:WorkWithSingleChartSeries

            //ExStart:ChartDataPoint 
            // Specifies whether by default the parent element shall inverts its colors if the value is negative.
            series0.InvertIfNegative = true;

            series0.Marker.Symbol = MarkerSymbol.Circle;
            series0.Marker.Size = 15;

            series1.Marker.Symbol = MarkerSymbol.Star;
            series1.Marker.Size = 10;
            //ExEnd:ChartDataPoint 

            doc.Save(ArtifactsDir + "WorkingWithCharts.SingleChartSeries.docx");
        }
    }
}
