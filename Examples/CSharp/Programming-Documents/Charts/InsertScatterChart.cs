using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Charts
{
    class InsertScatterChart
    {
        public static void Run()
        {
            //ExStart:InsertScatterChart
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithCharts();
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert Scatter chart.
            Shape shape = builder.InsertChart(ChartType.Scatter, 432, 252);
            Chart chart = shape.Chart;

            // Use this overload to add series to any type of Scatter charts.
            chart.Series.Add("AW Series 1", new double[] { 0.7, 1.8, 2.6 }, new double[] { 2.7, 3.2, 0.8 });

            dataDir = dataDir + "TestInsertScatterChart_out_.docx";
            doc.Save(dataDir);
            //ExEnd:InsertScatterChart
            Console.WriteLine("\nScatter chart created successfully.\nFile saved at " + dataDir);
        }        
    }
}
