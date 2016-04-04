' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_WorkingWithCharts()
Dim doc As New Document()
Dim builder As New DocumentBuilder(doc)

Dim shape As Shape = builder.InsertChart(ChartType.Bar, 432, 252)
Dim chart As Chart = shape.Chart

' Get first series.
Dim series0 As ChartSeries = shape.Chart.Series(0)
Dim dataLabelCollection As ChartDataLabelCollection = series0.DataLabels

' Add data label to the first and second point of the first series.
Dim chartDataLabel00 As ChartDataLabel = dataLabelCollection.Add(0)
Dim chartDataLabel01 As ChartDataLabel = dataLabelCollection.Add(1)

' Set properties.
chartDataLabel00.ShowLegendKey = True

' By default, when you add data labels to the data points in a pie chart, leader lines are displayed for data labels that are
' positioned far outside the end of data points. Leader lines create a visual connection between a data label and its 
' corresponding data point.
chartDataLabel00.ShowLeaderLines = True

chartDataLabel00.ShowCategoryName = False
chartDataLabel00.ShowPercentage = False
chartDataLabel00.ShowSeriesName = True
chartDataLabel00.ShowValue = True
chartDataLabel00.Separator = "/"
chartDataLabel01.ShowValue = True
dataDir = dataDir & Convert.ToString("SimpleBarChart_out_.docx")
doc.Save(dataDir)
