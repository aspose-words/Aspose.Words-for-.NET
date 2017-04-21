Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.Drawing.Charts
Public Class WorkWithSingleChartDataPoint
    Public Shared Sub Run()
        ' ExStart:WorkWithSingleChartDataPoint
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithCharts()
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        Dim shape As Shape = builder.InsertChart(ChartType.Line, 432, 252)
        Dim chart As Chart = shape.Chart

        ' Get first series.
        Dim series0 As ChartSeries = shape.Chart.Series(0)
        ' Get second series.
        Dim series1 As ChartSeries = shape.Chart.Series(1)
        Dim dataPointCollection As ChartDataPointCollection = series0.DataPoints

        ' Add data point to the first and second point of the first series.
        Dim dataPoint00 As ChartDataPoint = dataPointCollection.Add(0)
        Dim dataPoint01 As ChartDataPoint = dataPointCollection.Add(1)

        ' Set explosion.
        dataPoint00.Explosion = 50

        ' Set marker symbol and size.
        dataPoint00.Marker.Symbol = MarkerSymbol.Circle
        dataPoint00.Marker.Size = 15

        dataPoint01.Marker.Symbol = MarkerSymbol.Diamond
        dataPoint01.Marker.Size = 20

        ' Add data point to the third point of the second series.
        Dim dataPoint12 As ChartDataPoint = series1.DataPoints.Add(2)
        dataPoint12.InvertIfNegative = True
        dataPoint12.Marker.Symbol = MarkerSymbol.Star
        dataPoint12.Marker.Size = 20
        dataDir = dataDir & Convert.ToString("SingleChartDataPoint_out.docx")
        doc.Save(dataDir)
        ' ExEnd:WorkWithSingleChartDataPoint
        Console.WriteLine(Convert.ToString(vbLf & "Single line chart created successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
