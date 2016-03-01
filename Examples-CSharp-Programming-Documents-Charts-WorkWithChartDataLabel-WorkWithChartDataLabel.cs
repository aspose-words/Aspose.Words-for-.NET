// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithCharts();
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

Shape shape = builder.InsertChart(ChartType.Bar, 432, 252);
Chart chart = shape.Chart;

// Get first series.
ChartSeries series0 = shape.Chart.Series[0];
ChartDataLabelCollection dataLabelCollection = series0.DataLabels;

// Add data label to the first and second point of the first series.
ChartDataLabel chartDataLabel00 = dataLabelCollection.Add(0);
ChartDataLabel chartDataLabel01 = dataLabelCollection.Add(1);

// Set properties.
chartDataLabel00.ShowLegendKey = true;

// By default, when you add data labels to the data points in a pie chart, leader lines are displayed for data labels that are
// positioned far outside the end of data points. Leader lines create a visual connection between a data label and its 
// corresponding data point.
chartDataLabel00.ShowLeaderLines = true;

chartDataLabel00.ShowCategoryName = false;
chartDataLabel00.ShowPercentage = false;
chartDataLabel00.ShowSeriesName = true;
chartDataLabel00.ShowValue = true;
chartDataLabel00.Separator = "/";
chartDataLabel01.ShowValue = true;
dataDir = dataDir + @"SimpleBarChart_out_.docx";
doc.Save(dataDir);
