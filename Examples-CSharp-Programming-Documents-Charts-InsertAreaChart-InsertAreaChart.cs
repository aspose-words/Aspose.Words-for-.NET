// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithCharts();
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

// Insert Area chart.
Shape shape = builder.InsertChart(ChartType.Area, 432, 252);
Chart chart = shape.Chart;

// Use this overload to add series to any type of Area, Radar and Stock charts.
chart.Series.Add("AW Series 1", new DateTime[] { 
new DateTime(2002, 05, 01), 
new DateTime(2002, 06, 01),
new DateTime(2002, 07, 01),
new DateTime(2002, 08, 01),
new DateTime(2002, 09, 01)}, new double[] { 32, 32, 28, 12, 15 });
dataDir = dataDir + @"TestInsertAreaChart_out_.docx";
doc.Save(dataDir);
