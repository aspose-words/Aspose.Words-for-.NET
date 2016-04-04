// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

// Insert Column chart.
Shape shape = builder.InsertChart(ChartType.Column, 432, 252);
Chart chart = shape.Chart;

// Use this overload to add series to any type of Bar, Column, Line and Surface charts.
chart.Series.Add("AW Series 1", new string[] { "AW Category 1", "AW Category 2" }, new double[] { 1, 2 });

dataDir = dataDir + @"TestInsertChartColumn_out_.doc";
doc.Save(dataDir);
