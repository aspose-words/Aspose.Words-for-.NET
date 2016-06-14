using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Charts
{
    class InsertBubbleChart
    {
        public static void Run()
        {
            //ExStart:InsertBubbleChart
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithCharts();
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert Bubble chart.
            Shape shape = builder.InsertChart(ChartType.Bubble, 432, 252);
            Chart chart = shape.Chart;

            // Use this overload to add series to any type of Bubble charts.
            chart.Series.Add("AW Series 1", new double[] { 0.7, 1.8, 2.6 }, new double[] { 2.7, 3.2, 0.8 }, new double[] { 10, 4, 8 });
            dataDir = dataDir + @"TestInsertBubbleChart_out_.docx";
            doc.Save(dataDir);
            //ExEnd:InsertBubbleChart
            Console.WriteLine("\nBubble chart created successfully.\nFile saved at " + dataDir);
        }        
    }
}
