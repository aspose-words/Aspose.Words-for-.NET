// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithCharts();
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

Shape shape = builder.InsertChart(ChartType.Line, 432, 252);
Chart chart = shape.Chart;

// Get first series.
ChartSeries series0 = shape.Chart.Series[0];
// Get second series.
ChartSeries series1 = shape.Chart.Series[1];
ChartDataPointCollection dataPointCollection = series0.DataPoints;

// Add data point to the first and second point of the first series.
ChartDataPoint dataPoint00 = dataPointCollection.Add(0);
ChartDataPoint dataPoint01 = dataPointCollection.Add(1);

// Set explosion.
dataPoint00.Explosion = 50;

// Set marker symbol and size.
dataPoint00.Marker.Symbol = MarkerSymbol.Circle;
dataPoint00.Marker.Size = 15;

dataPoint01.Marker.Symbol = MarkerSymbol.Diamond;
dataPoint01.Marker.Size = 20;

// Add data point to the third point of the second series.
ChartDataPoint dataPoint12 = series1.DataPoints.Add(2);
dataPoint12.InvertIfNegative = true;
dataPoint12.Marker.Symbol = MarkerSymbol.Star;
dataPoint12.Marker.Size = 20;
dataDir = dataDir + @"SingleChartDataPoint_out_.docx";
doc.Save(dataDir);
