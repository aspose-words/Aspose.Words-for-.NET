Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.Drawing.Charts
Public Class InsertScatterChart
    Public Shared Sub Run()
        ' ExStart:InsertScatterChart
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithCharts()
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' Insert Scatter chart.
        Dim shape As Shape = builder.InsertChart(ChartType.Scatter, 432, 252)
        Dim chart As Chart = shape.Chart

        ' Use this overload to add series to any type of Scatter charts.
        chart.Series.Add("AW Series 1", New Double() {0.7, 1.8, 2.6}, New Double() {2.7, 3.2, 0.8})

        dataDir = dataDir & Convert.ToString("TestInsertScatterChart_out_.docx")
        doc.Save(dataDir)
        ' ExEnd:InsertScatterChart
        Console.WriteLine(Convert.ToString(vbLf & "Scatter chart created successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
