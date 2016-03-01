' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_WorkingWithCharts()
Dim doc As New Document()
Dim builder As New DocumentBuilder(doc)

' Insert Area chart.
Dim shape As Shape = builder.InsertChart(ChartType.Area, 432, 252)
Dim chart As Chart = shape.Chart

' Use this overload to add series to any type of Area, Radar and Stock charts.
chart.Series.Add("AW Series 1", New DateTime() {New DateTime(2002, 5, 1), New DateTime(2002, 6, 1), New DateTime(2002, 7, 1), New DateTime(2002, 8, 1), New DateTime(2002, 9, 1)}, New Double() {32, 32, 28, 12, 15})
dataDir = dataDir & Convert.ToString("TestInsertAreaChart_out_.docx")
doc.Save(dataDir)
