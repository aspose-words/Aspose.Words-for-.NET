Imports Microsoft.VisualBasic
Imports System.Drawing
Imports Aspose.Words
Imports Aspose.Words.Tables
Public Class ApplyStyle
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithTables()
        BuildTableWithStyle(dataDir)
        ExpandFormattingOnCellsAndRowFromStyle(dataDir)
    End Sub
    ''' <summary>
    ''' Shows how to build a new table with a table style applied.
    ''' </summary>
    Private Shared Sub BuildTableWithStyle(dataDir As String)
        'ExStart:BuildTableWithStyle
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        Dim table As Table = builder.StartTable()
        ' We must insert at least one row first before setting any table formatting.
        builder.InsertCell()
        ' Set the table style used based of the unique style identifier.
        ' Note that not all table styles are available when saving as .doc format.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1
        ' Apply which features should be formatted by the style.
        table.StyleOptions = TableStyleOptions.FirstColumn Or TableStyleOptions.RowBands Or TableStyleOptions.FirstRow
        table.AutoFit(AutoFitBehavior.AutoFitToContents)

        ' Continue with building the table as normal.
        builder.Writeln("Item")
        builder.CellFormat.RightPadding = 40
        builder.InsertCell()
        builder.Writeln("Quantity (kg)")
        builder.EndRow()

        builder.InsertCell()
        builder.Writeln("Apples")
        builder.InsertCell()
        builder.Writeln("20")
        builder.EndRow()

        builder.InsertCell()
        builder.Writeln("Bananas")
        builder.InsertCell()
        builder.Writeln("40")
        builder.EndRow()

        builder.InsertCell()
        builder.Writeln("Carrots")
        builder.InsertCell()
        builder.Writeln("50")
        builder.EndRow()

        dataDir = dataDir & Convert.ToString("DocumentBuilder.SetTableStyle_out_.docx")

        ' Save the document to disk.
        doc.Save(dataDir)
        'ExEnd:BuildTableWithStyle
        Console.WriteLine(Convert.ToString(vbLf & "Table created successfully with table style." & vbLf & "File saved at ") & dataDir)
    End Sub
    ''' <summary>
    ''' Shows how to expand the formatting from styles onto the rows and cells of the table as direct formatting.
    ''' </summary>
    Private Shared Sub ExpandFormattingOnCellsAndRowFromStyle(dataDir As String)
        'ExStart:ExpandFormattingOnCellsAndRowFromStyle
        Dim doc As New Document(dataDir & Convert.ToString("Table.TableStyle.docx"))

        ' Get the first cell of the first table in the document.
        Dim table As Table = DirectCast(doc.GetChild(NodeType.Table, 0, True), Table)
        Dim firstCell As Cell = table.FirstRow.FirstCell

        ' First print the color of the cell shading. This should be empty as the current shading
        ' is stored in the table style.
        Dim cellShadingBefore As Color = firstCell.CellFormat.Shading.BackgroundPatternColor
        Console.WriteLine("Cell shading before style expansion: " + cellShadingBefore.ToString())

        ' Expand table style formatting to direct formatting.
        doc.ExpandTableStylesToDirectFormatting()

        ' Now print the cell shading after expanding table styles. A blue background pattern color
        ' should have been applied from the table style.
        Dim cellShadingAfter As Color = firstCell.CellFormat.Shading.BackgroundPatternColor
        Console.WriteLine("Cell shading after style expansion: " + cellShadingAfter.ToString())
        'ExEnd:ExpandFormattingOnCellsAndRowFromStyle

    End Sub
End Class
