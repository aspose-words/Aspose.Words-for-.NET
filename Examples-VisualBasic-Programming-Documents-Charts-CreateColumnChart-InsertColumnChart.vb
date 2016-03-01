' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim doc As New Document()
Dim builder As New DocumentBuilder(doc)

' Insert Column chart.
Dim shape As Shape = builder.InsertChart(ChartType.Column, 432, 252)
Dim chart As Chart = shape.Chart

' Use this overload to add series to any type of Bar, Column, Line and Surface charts.
chart.Series.Add("AW Series 1", New String() {"AW Category 1", "AW Category 2"}, New Double() {1, 2})

dataDir = dataDir & Convert.ToString("TestInsertChartColumn_out_.doc")
doc.Save(dataDir)
