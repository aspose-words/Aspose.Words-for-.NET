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
            //ExFor:Charts.Chart
            //ExFor:Charts.Chart.Title
            //ExFor:Charts.ChartTitle
            //ExFor:Charts.ChartTitle.Overlay
            //ExFor:Charts.ChartTitle.Show
            //ExFor:Charts.ChartTitle.Text
            //ExSummary:Shows how to insert a chart and change its title.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use a document builder to insert a bar chart
            Shape chartShape = builder.InsertChart(ChartType.Bar, 400, 300);

            Assert.AreEqual(ShapeType.NonPrimitive, chartShape.ShapeType);
            Assert.True(chartShape.HasChart);

            // Get the chart object from the containing shape
            Chart chart = chartShape.Chart;
            
            // Set the title text, which appears at the top center of the chart and modify its appearance
            ChartTitle title = chart.Title;
            title.Text = "MyChart";
            title.Overlay = true;
            title.Show = true;

            doc.Save(ArtifactsDir + "Charts.ChartSeries.docx");
            //ExEnd
        }

        //ExStart
        //ExFor:Charts.ChartSeries
        //ExFor:Charts.ChartSeries.DataLabels
        //ExFor:Charts.ChartSeries.DataPoints
        //ExFor:Charts.ChartSeries.Name
        //ExFor:Charts.ChartSeries.Smooth
        //ExFor:Charts.ChartDataLabel
        //ExFor:Charts.ChartDataLabel.Index
        //ExFor:Charts.ChartDataLabel.IsVisible
        //ExFor:Charts.ChartDataLabel.NumberFormat
        //ExFor:Charts.ChartDataLabel.Separator
        //ExFor:Charts.ChartDataLabel.ShowBubbleSize
        //ExFor:Charts.ChartDataLabel.ShowCategoryName
        //ExFor:Charts.ChartDataLabel.ShowDataLabelsRange
        //ExFor:Charts.ChartDataLabel.ShowLeaderLines
        //ExFor:Charts.ChartDataLabel.ShowLegendKey
        //ExFor:Charts.ChartDataLabel.ShowPercentage
        //ExFor:Charts.ChartDataLabel.ShowSeriesName
        //ExFor:Charts.ChartDataLabel.ShowValue
        //ExFor:Charts.ChartDataLabelCollection
        //ExFor:Charts.ChartDataLabelCollection.Add(System.Int32)
        //ExFor:Charts.ChartDataLabelCollection.Clear
        //ExFor:Charts.ChartDataLabelCollection.Count
        //ExFor:Charts.ChartDataLabelCollection.GetEnumerator
        //ExFor:Charts.ChartDataLabelCollection.Item(System.Int32)
        //ExFor:Charts.ChartDataLabelCollection.RemoveAt(System.Int32)
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

            // The line for this series will be smoothed
            chart.Series[0].Smooth = true;

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
            chart.Series[1].DataLabels.RemoveAt(2);

            // We can also clear an entire data label collection for one whole series
            chart.Series[2].DataLabels.Clear();

            doc.Save(ArtifactsDir + "Charts.ChartDataLabels.docx");
        }

        /// <summary>
        /// Apply uniform data labels to a number (determined by labelsCount) of data points in a series
        /// Also apply a number format and separator
        /// </summary>
        private void ApplyDataLabels(ChartSeries series, int labelsCount, string numberFormat, string separator)
        {
            for (int i = 0; i < labelsCount; i++)
            {
                ChartDataLabel label = series.DataLabels.Add(i);
                Assert.False(label.IsVisible);

                // Edit the appearance of the new data label
                label.ShowCategoryName = true;
                label.ShowSeriesName = true;
                label.ShowValue = true;
                label.ShowLeaderLines = true;
                label.ShowLegendKey = true;
                label.ShowPercentage = false;
                Assert.False(label.ShowDataLabelsRange);

                // Apply number format and separator
                label.NumberFormat.FormatCode = numberFormat;
                label.Separator = separator;

                // Attributes that apply exclusively to bubble graphs
                Assert.False(series.Bubble3D);
                Assert.False(label.ShowBubbleSize);

                // The label automatically becomes visible
                Assert.True(label.IsVisible);
            }
        }
        //ExEnd

        //ExStart
        //ExFor:Charts.Chart.Series
        //ExFor:Charts.ChartSeriesCollection.Add(String,DateTime[],Double[])
        //ExFor:Charts.ChartSeriesCollection.Add(String,Double[],Double[])
        //ExFor:Charts.ChartSeriesCollection.Add(String,Double[],Double[],Double[])
        //ExFor:Charts.ChartSeriesCollection.Add(String,String[],Double[])
        //ExFor:Charts.ChartType
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

            // Create and name 3 categories with a string array
            string[] categories = { "Category 1", "Category 2", "Category 3" };

            // Create 2 series of data, each with one point for every category
            // This will generate a column graph with 3 clusters of 2 bars
            chart.Series.Add("Series 1", categories, new [] { 76.6, 82.1, 91.6 });
            chart.Series.Add("Series 2", categories, new [] { 64.2, 79.5, 94.0 });

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

            // 4: Each series will be built from three data arrays, used for bubble charts
            chart = AppendChart(builder, ChartType.Bubble, 300, 300);

            // The first two arrays contain X/Y coordinates like above and the third determines the thickness of each point
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
            chart.Series.Add("Series 1", new[] { 1.0, 2.0, 3.0, 4.0, 5.0 }, new[] { 1.0, 20.0, 400.0, 8000.0, 160000.0 });

            // The scaling of the X axis is linear by default, which means it will display "0, 1, 2, 3..."
            Assert.AreEqual(AxisScaleType.Linear, chart.AxisX.Scaling.Type);

            // Linear axis scaling is suitable for our X-values, but not our erratic Y-values 
            // We can set the scaling of the Y-axis to Logarithmic with a base of 20
            // The Y-axis will now display "1, 20, 400, 8000...", which is ideal for accurate representation of this set of Y-values
            chart.AxisY.Scaling.Type = AxisScaleType.Logarithmic;
            chart.AxisY.Scaling.LogBase = 20.0;

            doc.Save(ArtifactsDir + "Charts.AxisScaling.docx");
            //ExEnd
        }

        [Test]
        public void AxisBound()
        {
            //ExStart
            //ExFor:Charts.AxisBound.#ctor
            //ExFor:Charts.AxisBound.IsAuto
            //ExFor:Charts.AxisBound.Value
            //ExFor:Charts.AxisBound.ValueAsDate
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
        }

        [Test]
        public void ChartLegend()
        {
            //ExStart
            //ExFor:Charts.Chart.Legend
            //ExFor:Charts.ChartLegend
            //ExFor:Charts.ChartLegend.Overlay
            //ExFor:Charts.ChartLegend.Position
            //ExSummary:Shows how to edit the appearance of a chart's legend.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a line graph
            Shape chartShape = builder.InsertChart(ChartType.Line, 450, 300);
            Chart chart = chartShape.Chart;

            // Get its legend
            ChartLegend legend = chart.Legend;

            // By default, other elements of a chart will not overlap with its legend
            Assert.False(legend.Overlay);

            // We can move its position by setting this attribute
            legend.Position = LegendPosition.TopRight;

            doc.Save(ArtifactsDir + "Charts.ChartLegend.docx");
            //ExEnd
        }
    }
}