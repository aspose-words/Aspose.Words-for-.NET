' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim doc As New Document()
Dim builder As New DocumentBuilder(doc)

' Add chart with default data. You can specify different chart types and sizes.
Dim shape As Shape = builder.InsertChart(ChartType.Column, 432, 252)

' Chart property of Shape contains all chart related options.
Dim chart As Chart = shape.Chart

' Get chart series collection.
Dim seriesColl As ChartSeriesCollection = chart.Series
' Check series count.
Console.WriteLine(seriesColl.Count)

' Delete default generated series.
seriesColl.Clear()

' Create category names array, in this example we have two categories.
Dim categories As String() = New String() {"AW Category 1", "AW Category 2"}

' Adding new series. Please note, data arrays must not be empty and arrays must be the same size.
seriesColl.Add("AW Series 1", categories, New Double() {1, 2})
seriesColl.Add("AW Series 2", categories, New Double() {3, 4})
seriesColl.Add("AW Series 3", categories, New Double() {5, 6})
seriesColl.Add("AW Series 4", categories, New Double() {7, 8})
seriesColl.Add("AW Series 5", categories, New Double() {9, 10})

dataDir = dataDir & Convert.ToString("TestInsertSimpleChartColumn_out_.doc")
doc.Save(dataDir)
