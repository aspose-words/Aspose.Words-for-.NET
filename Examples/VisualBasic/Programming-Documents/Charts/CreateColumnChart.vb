Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.Drawing.Charts
Public Class CreateColumnChart
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithCharts()
        InsertSimpleColumnChart(dataDir)
        InsertColumnChart(dataDir)
    End Sub
    ''' <summary>
    '''  Shows how to insert a simple column chart into the document using DocumentBuilder.InsertChart method.
    ''' </summary>             
    Private Shared Sub InsertSimpleColumnChart(dataDir As String)
        ' ExStart:InsertSimpleColumnChart
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' Add chart with default data. You can specify different chart types and sizes.
        Dim shape As Shape = builder.InsertChart(ChartType.Column, 432, 252)

        ' Chart property of Shape contains all chart related options.
        Dim chart As Chart = shape.Chart

        ' ExStart:ChartSeriesCollection 
        ' Get chart series collection.
        Dim seriesColl As ChartSeriesCollection = chart.Series
        ' Check series count.
        Console.WriteLine(seriesColl.Count)
        ' ExEnd:ChartSeriesCollection 

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
        ' ExEnd:InsertSimpleColumnChart
        Console.WriteLine(Convert.ToString(vbLf & "Simple column chart created successfully." & vbLf & "File saved at ") & dataDir)

    End Sub
    ''' <summary>
    '''  Shows how to insert a column chart into the document using DocumentBuilder.InsertChart method.
    ''' </summary>             
    Private Shared Sub InsertColumnChart(dataDir As String)
        ' ExStart:InsertColumnChart
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' Insert Column chart.
        Dim shape As Shape = builder.InsertChart(ChartType.Column, 432, 252)
        Dim chart As Chart = shape.Chart

        ' Use this overload to add series to any type of Bar, Column, Line and Surface charts.
        chart.Series.Add("AW Series 1", New String() {"AW Category 1", "AW Category 2"}, New Double() {1, 2})

        dataDir = dataDir & Convert.ToString("TestInsertChartColumn_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:InsertColumnChart
        Console.WriteLine(Convert.ToString(vbLf & "Column chart created successfully." & vbLf & "File saved at ") & dataDir)

    End Sub
End Class
