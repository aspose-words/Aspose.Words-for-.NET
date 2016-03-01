Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.Drawing.Charts
Public Class WorkWithSingleChartSeries
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithCharts()
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        Dim shape As Shape = builder.InsertChart(ChartType.Line, 432, 252)
        Dim chart As Chart = shape.Chart
        ' ExStart:WorkWithSingleChartSeries
        ' Get first series.
        Dim series0 As ChartSeries = shape.Chart.Series(0)

        ' Get second series.
        Dim series1 As ChartSeries = shape.Chart.Series(1)

        ' Change first series name.
        series0.Name = "My Name1"

        ' Change second series name.
        series1.Name = "My Name2"

        ' You can also specify whether the line connecting the points on the chart shall be smoothed using Catmull-Rom splines.
        series0.Smooth = True
        series1.Smooth = True
        ' ExEnd:WorkWithSingleChartSeries
        ' ExStart:ChartDataPoint 
        ' Specifies whether by default the parent element shall inverts its colors if the value is negative.
        series0.InvertIfNegative = True

        ' Set default marker symbol and size.
        series0.Marker.Symbol = MarkerSymbol.Circle
        series0.Marker.Size = 15

        series1.Marker.Symbol = MarkerSymbol.Star
        series1.Marker.Size = 10
        ' ExEnd:ChartDataPoint 
        dataDir = dataDir & Convert.ToString("SingleChartSeries_out_.docx")
        doc.Save(dataDir)

        Console.WriteLine(Convert.ToString(vbLf & "Chart created successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
