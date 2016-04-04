// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithCharts();
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

Shape shape = builder.InsertChart(ChartType.Line, 432, 252);
Chart chart = shape.Chart;

// Determines whether the title shall be shown for this chart. Default is true.
chart.Title.Show = true;

// Setting chart Title.
chart.Title.Text = "Sample Line Chart Title";

// Determines whether other chart elements shall be allowed to overlap title.
chart.Title.Overlay = false;

// Please note if null or empty value is specified as title text, auto generated title will be shown.

// Determines how legend shall be shown for this chart.
chart.Legend.Position = LegendPosition.Left;
chart.Legend.Overlay = true;
dataDir = dataDir + @"SimpleLineChart_out_.docx";
doc.Save(dataDir);
