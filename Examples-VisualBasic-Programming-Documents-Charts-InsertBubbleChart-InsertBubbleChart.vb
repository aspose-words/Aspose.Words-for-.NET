' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_WorkingWithCharts()
Dim doc As New Document()
Dim builder As New DocumentBuilder(doc)

' Insert Bubble chart.
Dim shape As Shape = builder.InsertChart(ChartType.Bubble, 432, 252)
Dim chart As Chart = shape.Chart

' Use this overload to add series to any type of Bubble charts.
chart.Series.Add("AW Series 1", New Double() {0.7, 1.8, 2.6}, New Double() {2.7, 3.2, 0.8}, New Double() {10, 4, 8})
dataDir = dataDir & Convert.ToString("TestInsertBubbleChart_out_.docx")
doc.Save(dataDir)
