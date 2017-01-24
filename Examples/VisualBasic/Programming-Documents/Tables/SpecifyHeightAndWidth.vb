Imports Microsoft.VisualBasic
Imports System.Drawing
Imports Aspose.Words
Imports Aspose.Words.Tables
Public Class SpecifyHeightAndWidth
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithTables()
        AutoFitToPageWidth(dataDir)
        SetPreferredWidthSettings(dataDir)
        RetrievePreferredWidthType(dataDir)
    End Sub
    ''' <summary>
    ''' Shows how to set a table to auto fit to 50% of the page width.
    ''' </summary>
    Private Shared Sub AutoFitToPageWidth(dataDir As String)
        ' ExStart:AutoFitToPageWidth
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' Insert a table with a width that takes up half the page width.
        Dim table As Table = builder.StartTable()

        ' Insert a few cells
        builder.InsertCell()
        table.PreferredWidth = PreferredWidth.FromPercent(50)
        builder.Writeln("Cell #1")

        builder.InsertCell()
        builder.Writeln("Cell #2")

        builder.InsertCell()
        builder.Writeln("Cell #3")

        dataDir = dataDir & Convert.ToString("Table.PreferredWidth_out.doc")

        ' Save the document to disk.
        doc.Save(dataDir)
        ' ExEnd:AutoFitToPageWidth
        Console.WriteLine(Convert.ToString(vbLf & "Table autofit successfully to 50% of the page width." & vbLf & "File saved at ") & dataDir)
    End Sub
    ''' <summary>
    ''' Shows how to set the different preferred width settings.
    ''' </summary>
    Private Shared Sub SetPreferredWidthSettings(dataDir As String)
        ' ExStart:SetPreferredWidthSettings
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' Insert a table row made up of three cells which have different preferred widths.
        Dim table As Table = builder.StartTable()

        ' Insert an absolute sized cell.
        builder.InsertCell()
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(40)
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightYellow
        builder.Writeln("Cell at 40 points width")

        ' Insert a relative (percent) sized cell.
        builder.InsertCell()
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(20)
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue
        builder.Writeln("Cell at 20% width")

        ' Insert a auto sized cell.
        builder.InsertCell()
        builder.CellFormat.PreferredWidth = PreferredWidth.Auto
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen
        builder.Writeln("Cell automatically sized. The size of this cell is calculated from the table preferred width.")
        builder.Writeln("In this case the cell will fill up the rest of the available space.")

        dataDir = dataDir & Convert.ToString("Table.CellPreferredWidths_out.doc")
        ' Save the document to disk.
        doc.Save(dataDir)
        ' ExEnd:SetPreferredWidthSettings
        Console.WriteLine(Convert.ToString(vbLf & "Different preferred width settings set successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
    ''' <summary>
    ''' Shows how to retrieves the preferred width type of a table cell.
    ''' </summary>
    Private Shared Sub RetrievePreferredWidthType(dataDir As String)
        ' ExStart:RetrievePreferredWidthType
        Dim doc As New Document(dataDir & Convert.ToString("Table.SimpleTable.doc"))

        ' Retrieve the first table in the document.
        Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)
        ' ExStart:AllowAutoFit
        table.AllowAutoFit = True
        ' ExEnd:AllowAutoFit

        Dim firstCell As Cell = table.FirstRow.FirstCell
        Dim type As PreferredWidthType = firstCell.CellFormat.PreferredWidth.Type
        Dim value As Double = firstCell.CellFormat.PreferredWidth.Value

        ' ExEnd:RetrievePreferredWidthType
        Console.WriteLine(vbLf & "Table preferred width type value is " + value.ToString())
    End Sub
End Class
