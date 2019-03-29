using System;
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
        public void ChartSeries()
        {
            //ExStart
            //ExFor:Charts.Chart.Title
            //ExFor:Charts.ChartTitle
            //ExFor:Charts.ChartTitle.Overlay
            //ExFor:Charts.ChartTitle.Show
            //ExFor:Charts.ChartTitle.Text
            //ExSumary:Shows how to insert a chart and change its title.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Shape chartShape = builder.InsertChart(ChartType.Bar, 400, 300);

            Assert.AreEqual(ShapeType.NonPrimitive, chartShape.ShapeType);
            Assert.True(chartShape.HasChart);

            Chart chart = chartShape.Chart;
            
            ChartTitle title = chart.Title;
            title.Text = "MyChart";
            title.Overlay = true;
            title.Show = true;

            doc.Save(ArtifactsDir + "Charts.BarChart.docx");
            //ExEnd
        }

        //ExStart
        //ExFor:Charts.Chart.Series
        //ExFor:Charts.ChartSeriesCollection.Add(String,DateTime[],Double[])
        //ExFor:Charts.ChartSeriesCollection.Add(String,Double[],Double[])
        //ExFor:Charts.ChartSeriesCollection.Add(String,Double[],Double[],Double[])
        //ExFor:Charts.ChartSeriesCollection.Add(String,String[],Double[])
        //ExSummary:Shows an appropriate graph type for each chart series.
        [Test] //ExSkip
        public void ChartSeriesCollection()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            // There are 4 ways of populating a chart's series collection
            // 1: Each series has a string array of categories, each with a corresponding data value
            // Some of the other possible applications are bar, column, line and surface charts
            Chart chart = AppendChart(builder, ChartType.Column, 300, 300);

            // The number of categories and their names are defined by a string array
            // Each category will be represented by a bar cluster in the chart
            string[] categories = { "Category 1", "Category 2", "Category 3" };

            // The chart will have 3 bar clusters and each cluster will have two columns to compare,
            // side by side, data values of the index of the cluster's category from each series
            chart.Series.Add("Series 1", categories, new [] { 76.6, 82.1, 91.6 });
            chart.Series.Add("Series 2", categories, new [] { 64.2, 79.5, 94.0 });

            // 2: Each series will have a collection of dates with a corresponding value for each date
            // Area, radar and stock charts are some of the appropriate chart types for this
            chart = AppendChart(builder, ChartType.Area, 300, 300);

            // Create a collection of dates
            DateTime[] dates = { new DateTime(2014, 3, 31),
                new DateTime(2017, 1, 23),
                new DateTime(2017, 6, 18),
                new DateTime(2019, 11, 22),
                new DateTime(2020, 9, 7)
            };

            // Add a series with the collection of dates and an array of data, with a value for each date
            // Our area chart will adjust for the sporadic DateTime values and display the data along a linear time scale
            chart.Series.Add("Series 1", dates, new [] { 15.8, 21.5, 22.9, 28.7, 33.1 });

            // 3: Each series will take two data arrays
            // Appropriate for scatter plots
            chart = AppendChart(builder, ChartType.Scatter, 300, 300);

            // Two arrays are passed, both of the same size
            // The first array contains X-coordinates and the second contains respective Y-coordinates of the points that will populate the scatter chart
            // Each series adds 8 points which will share the same color, to differentiate from the points of another series
            chart.Series.Add("Series 1", new[] { 3.1, 3.5, 6.3, 4.1, 2.2, 8.3, 1.2, 3.6 }, new[] { 3.1, 6.3, 4.6, 0.9, 8.5, 4.2, 2.3, 9.9 });
            chart.Series.Add("Series 2", new[] { 2.6, 7.3, 4.5, 6.6, 2.1, 9.3, 0.7, 3.3 }, new[] { 7.1, 6.6, 3.5, 7.8, 7.7, 9.5, 1.3, 4.6 });

            // 4: Each series will be built from three data arrays, used for bubble charts
            chart = AppendChart(builder, ChartType.Bubble, 300, 300);

            // The first two passed data arrays are X and Y coordinates, like in the scatter plot above
            // The third data array specifies the displayed thickness of each point
            chart.Series.Add("Series 1", new [] { 1.1, 5.0, 9.8 }, new [] { 1.2, 4.9, 9.9 }, new [] { 2.0, 4.0, 8.0 });

            doc.Save(ArtifactsDir + "Charts.ChartSeriesCollection.docx");
        }

        /// <summary>
        /// Get the DocumentBuilder to insert a chart of a specified ChartType, width and height and clean out its default data
        /// </summary>
        private Chart AppendChart(DocumentBuilder builder, ChartType chartType, double width, double height)
        {
            Shape chartShape = builder.InsertChart(chartType, width, height);
            Chart chart = chartShape.Chart;
            chart.Series.Clear();

            Assert.AreEqual(0, chart.Series.Count);

            return chart;
        }
        //ExEnd

        [Test]
        public void AxisScaling()
        {
            //ExStart
            //ExFor:Charts.AxisScaleType
            //ExFor:Charts.AxisScaling
            //ExFor:Charts.AxisScaling.LogBase
            //ExFor:Charts.AxisScaling.Type
            //ExSummary:Shows how to set up logarithmic axis scaling.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a scatter chart and clear its default data series
            Shape chartShape = builder.InsertChart(ChartType.Scatter, 450, 300);
            Chart chart = chartShape.Chart;
            chart.Series.Clear();

            // Insert a series with X/Y coordinates for 5 points
            // Note that, because of the erratic Y values, a graph with linear Y-axis scaling may not produce a satisfactory result
            chart.Series.Add("Series 1", new[] { 1.0, 2.0, 3.0, 4.0, 5.0 }, new[] { 1.0, 10.0, 100.0, 1000.0, 10000.0 });

            // The scaling of the X axis is linear,
            // which means that the X value goes up by the same amount (in this case, 1) at each vertical line
            // This X-axis ruler will display "1  2  3.."
            Assert.AreEqual(AxisScaleType.Linear, chart.AxisX.Scaling.Type);

            // As for the Y axis, we can set the scaling to Logarithmic,
            // which means that the Y-values go up by an order of magnitude at each horizontal line
            // The order of magnitude is decided by the LogBase attribute, which we will leave at 10, its default value
            // This Y-axis ruler will display "1  10  100..."
            chart.AxisY.Scaling.Type = AxisScaleType.Logarithmic;
            Assert.AreEqual(10.0, chart.AxisX.Scaling.LogBase);

            doc.Save(ArtifactsDir + "Charts.AxisScaling.docx");
            //ExEnd
        }
    }
}