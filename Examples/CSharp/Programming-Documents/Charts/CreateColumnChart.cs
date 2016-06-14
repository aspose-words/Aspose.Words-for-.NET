using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Charts
{
    class CreateColumnChart
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithCharts();
            InsertSimpleColumnChart(dataDir);
            InsertColumnChart(dataDir);            
        }
        /// <summary>
        ///  Shows how to insert a simple column chart into the document using DocumentBuilder.InsertChart method.
        /// </summary>             
        private static void InsertSimpleColumnChart(string dataDir)
        {
            //ExStart:InsertSimpleColumnChart
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add chart with default data. You can specify different chart types and sizes.
            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);

            // Chart property of Shape contains all chart related options.
            Chart chart = shape.Chart;

            //ExStart:ChartSeriesCollection 
            // Get chart series collection.
            ChartSeriesCollection seriesColl = chart.Series;
            // Check series count.
            Console.WriteLine(seriesColl.Count);
            //ExEnd:ChartSeriesCollection 

            // Delete default generated series.
            seriesColl.Clear();

            // Create category names array, in this example we have two categories.
            string[] categories = new string[] { "AW Category 1", "AW Category 2" };

            // Adding new series. Please note, data arrays must not be empty and arrays must be the same size.
            seriesColl.Add("AW Series 1", categories, new double[] { 1, 2 });
            seriesColl.Add("AW Series 2", categories, new double[] { 3, 4 });
            seriesColl.Add("AW Series 3", categories, new double[] { 5, 6 });
            seriesColl.Add("AW Series 4", categories, new double[] { 7, 8 });
            seriesColl.Add("AW Series 5", categories, new double[] { 9, 10 });

            dataDir = dataDir + @"TestInsertSimpleChartColumn_out_.doc";
            doc.Save(dataDir);
            //ExEnd:InsertSimpleColumnChart
            Console.WriteLine("\nSimple column chart created successfully.\nFile saved at " + dataDir);
            
        }
        /// <summary>
        ///  Shows how to insert a column chart into the document using DocumentBuilder.InsertChart method.
        /// </summary>             
        private static void InsertColumnChart(string dataDir)
        {
            //ExStart:InsertColumnChart
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert Column chart.
            Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
            Chart chart = shape.Chart;

            // Use this overload to add series to any type of Bar, Column, Line and Surface charts.
            chart.Series.Add("AW Series 1", new string[] { "AW Category 1", "AW Category 2" }, new double[] { 1, 2 });

            dataDir = dataDir + @"TestInsertChartColumn_out_.doc";
            doc.Save(dataDir);
            //ExEnd:InsertColumnChart
            Console.WriteLine("\nColumn chart created successfully.\nFile saved at " + dataDir);

        }
    }
}
