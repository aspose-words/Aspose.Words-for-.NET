Imports System.IO
Imports Aspose.Words
Imports System.Drawing
Imports Aspose.Words.Tables

Class DocumentBuilderBuildTable
    Public Shared Sub Run()
        ' ExStart:DocumentBuilderBuildTable
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        ' Initialize document.
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        Dim table As Table = builder.StartTable()

        ' Insert a cell
        builder.InsertCell()
        ' Use fixed column widths.
        table.AutoFit(AutoFitBehavior.FixedColumnWidths)

        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center
        builder.Write("This is row 1 cell 1")

        ' Insert a cell
        builder.InsertCell()
        builder.Write("This is row 1 cell 2")

        builder.EndRow()

        ' Insert a cell
        builder.InsertCell()

        ' Apply new row formatting
        builder.RowFormat.Height = 100
        builder.RowFormat.HeightRule = HeightRule.Exactly

        builder.CellFormat.Orientation = TextOrientation.Upward
        builder.Writeln("This is row 2 cell 1")

        ' Insert a cell
        builder.InsertCell()
        builder.CellFormat.Orientation = TextOrientation.Downward
        builder.Writeln("This is row 2 cell 2")

        builder.EndRow()

        builder.EndTable()
        dataDir = dataDir & Convert.ToString("DocumentBuilderBuildTable_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderBuildTable
        Console.WriteLine(Convert.ToString(vbLf & "Table build successfully using DocumentBuilder." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
