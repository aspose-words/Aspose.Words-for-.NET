Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.Drawing.Charts
Public Class CreateChartUsingShape
    Public Shared Sub Run()
        ' ExStart:CreateChartUsingShape
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithCharts()
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        Dim shape As Shape = builder.InsertChart(ChartType.Line, 432, 252)
        Dim chart As Chart = shape.Chart

        ' Determines whether the title shall be shown for this chart. Default is true.
        chart.Title.Show = True

        ' Setting chart Title.
        chart.Title.Text = "Sample Line Chart Title"

        ' Determines whether other chart elements shall be allowed to overlap title.
        chart.Title.Overlay = False

        ' Please note if null or empty value is specified as title text, auto generated title will be shown.

        ' Determines how legend shall be shown for this chart.
        chart.Legend.Position = LegendPosition.Left
        chart.Legend.Overlay = True
        dataDir = dataDir & Convert.ToString("SimpleLineChart_out_.docx")
        doc.Save(dataDir)
        ' ExEnd:CreateChartUsingShape
        Console.WriteLine(Convert.ToString(vbLf & "Simple line chart created successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
