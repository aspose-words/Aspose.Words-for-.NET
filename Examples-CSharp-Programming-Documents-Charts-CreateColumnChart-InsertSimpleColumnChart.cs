// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

// Add chart with default data. You can specify different chart types and sizes.
Shape shape = builder.InsertChart(ChartType.Column, 432, 252);

// Chart property of Shape contains all chart related options.
Chart chart = shape.Chart;

// Get chart series collection.
ChartSeriesCollection seriesColl = chart.Series;
// Check series count.
Console.WriteLine(seriesColl.Count);

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
