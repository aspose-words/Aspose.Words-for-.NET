Imports Microsoft.VisualBasic
Imports System.Drawing
Imports Aspose.Words
Imports Aspose.Words.Tables
Public Class ApplyFormatting
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithTables()
        ApplyOutlineBorder(dataDir)
        BuildTableWithBordersEnabled(dataDir)
        ModifyRowFormatting(dataDir)
        ApplyRowFormatting(dataDir)
        ModifyCellFormatting(dataDir)
        FormatTableAndCellWithDifferentBorders(dataDir)
    End Sub
    ''' <summary>
    ''' Shows how to create a table that contains a single cell and apply row formatting.
    ''' </summary>
    Private Shared Sub ApplyRowFormatting(dataDir As String)
        'ExStart:ApplyRowFormatting
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        Dim table As Table = builder.StartTable()
        builder.InsertCell()

        ' Set the row formatting
        Dim rowFormat As RowFormat = builder.RowFormat
        rowFormat.Height = 100
        rowFormat.HeightRule = HeightRule.Exactly
        ' These formatting properties are set on the table and are applied to all rows in the table.
        table.LeftPadding = 30
        table.RightPadding = 30
        table.TopPadding = 30
        table.BottomPadding = 30

        builder.Writeln("I'm a wonderful formatted row.")

        builder.EndRow()
        builder.EndTable()

        dataDir = dataDir & Convert.ToString("Table.ApplyRowFormatting_out_.doc")

        ' Save the document to disk.
        doc.Save(dataDir)
        'ExEnd:ApplyRowFormatting
        Console.WriteLine(Convert.ToString(vbLf & "Row formatting applied successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
    Private Shared Sub ApplyOutlineBorder(dataDir As String)
        ' ExStart:ApplyOutlineBorder
        Dim doc As New Document(dataDir & Convert.ToString("Table.EmptyTable.doc"))

        Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)
        ' Align the table to the center of the page.
        table.Alignment = TableAlignment.Center
        ' Clear any existing borders from the table.
        table.ClearBorders()

        ' Set a green border around the table but not inside. 
        table.SetBorder(BorderType.Left, LineStyle.[Single], 1.5, Color.Green, True)
        table.SetBorder(BorderType.Right, LineStyle.[Single], 1.5, Color.Green, True)
        table.SetBorder(BorderType.Top, LineStyle.[Single], 1.5, Color.Green, True)
        table.SetBorder(BorderType.Bottom, LineStyle.[Single], 1.5, Color.Green, True)

        ' Fill the cells with a light green solid color.
        table.SetShading(TextureIndex.TextureSolid, Color.LightGreen, Color.Empty)
        dataDir = dataDir & Convert.ToString("Table.SetOutlineBorders_out_.doc")
        ' Save the document to disk.
        doc.Save(dataDir)
        ' ExEnd:ApplyOutlineBorder
        Console.WriteLine(Convert.ToString(vbLf & "Outline border applied successfully to a table." & vbLf & "File saved at ") & dataDir)
    End Sub
    Private Shared Sub BuildTableWithBordersEnabled(dataDir As String)
        ' ExStart:BuildTableWithBordersEnabled
        Dim doc As New Document(dataDir & Convert.ToString("Table.EmptyTable.doc"))

        Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)
        ' Clear any existing borders from the table.
        table.ClearBorders()
        ' Set a green border around and inside the table.
        table.SetBorders(LineStyle.[Single], 1.5, Color.Green)

        dataDir = dataDir & Convert.ToString("Table.SetAllBorders_out_.doc")
        ' Save the document to disk.
        doc.Save(dataDir)
        ' ExEnd:BuildTableWithBordersEnabled
        Console.WriteLine(Convert.ToString(vbLf & "Table build successfully with all borders enabled." & vbLf & "File saved at ") & dataDir)
    End Sub
    Private Shared Sub ModifyRowFormatting(dataDir As String)
        ' ExStart:ModifyRowFormatting
        Dim doc As New Document(dataDir & Convert.ToString("Table.Document.doc"))
        Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)

        ' Retrieve the first row in the table.
        Dim firstRow As Row = table.FirstRow
        ' Modify some row level properties.
        firstRow.RowFormat.Borders.LineStyle = LineStyle.None
        firstRow.RowFormat.HeightRule = HeightRule.Auto
        firstRow.RowFormat.AllowBreakAcrossPages = True
        ' ExEnd:ModifyRowFormatting
        Console.WriteLine(vbLf & "Some row level properties modified successfully.")
    End Sub
    Private Shared Sub ModifyCellFormatting(dataDir As String)
        ' ExStart:ModifyCellFormatting
        Dim doc As New Document(dataDir & Convert.ToString("Table.Document.doc"))
        Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)

        ' Retrieve the first cell in the table.
        Dim firstCell As Cell = table.FirstRow.FirstCell
        ' Modify some cell level properties.
        firstCell.CellFormat.Width = 30
        ' in points
        firstCell.CellFormat.Orientation = TextOrientation.Downward
        firstCell.CellFormat.Shading.ForegroundPatternColor = Color.LightGreen
        ' ExEnd:ModifyCellFormatting
        Console.WriteLine(vbLf & "Some cell level properties modified successfully.")
    End Sub
    Private Shared Sub FormatTableAndCellWithDifferentBorders(dataDir As String)
        ' ExStart:FormatTableAndCellWithDifferentBorders
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        Dim table As Table = builder.StartTable()
        builder.InsertCell()

        ' Set the borders for the entire table.
        table.SetBorders(LineStyle.[Single], 2.0, Color.Black)
        ' Set the cell shading for this cell.
        builder.CellFormat.Shading.BackgroundPatternColor = Color.Red
        builder.Writeln("Cell #1")

        builder.InsertCell()
        ' Specify a different cell shading for the second cell.
        builder.CellFormat.Shading.BackgroundPatternColor = Color.Green
        builder.Writeln("Cell #2")

        ' End this row.
        builder.EndRow()

        ' Clear the cell formatting from previous operations.
        builder.CellFormat.ClearFormatting()

        ' Create the second row.
        builder.InsertCell()

        ' Create larger borders for the first cell of this row. This will be different.
        ' compared to the borders set for the table.
        builder.CellFormat.Borders.Left.LineWidth = 4.0
        builder.CellFormat.Borders.Right.LineWidth = 4.0
        builder.CellFormat.Borders.Top.LineWidth = 4.0
        builder.CellFormat.Borders.Bottom.LineWidth = 4.0
        builder.Writeln("Cell #3")

        builder.InsertCell()
        ' Clear the cell formatting from the previous cell.
        builder.CellFormat.ClearFormatting()
        builder.Writeln("Cell #4")
        ' Save finished document.
        doc.Save(dataDir & Convert.ToString("Table.SetBordersAndShading_out_.doc"))
        ' ExEnd:FormatTableAndCellWithDifferentBorders
        Console.WriteLine(Convert.ToString(vbLf & "format table and cell with different borders and shadings successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
