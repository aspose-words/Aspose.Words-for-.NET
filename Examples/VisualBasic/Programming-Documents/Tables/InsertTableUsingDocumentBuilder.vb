Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Drawing
Imports System.Diagnostics
Imports Aspose.Words
Imports Aspose.Words.Tables
Public Class InsertTableUsingDocumentBuilder
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithTables()

        SimpleTable(dataDir)
        FormattedTable(dataDir)
        NestedTable(dataDir)
    End Sub
    Private Shared Sub SimpleTable(dataDir As String)
        ' ExStart:SimpleTable
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)
        ' We call this method to start building the table.
        builder.StartTable()
        builder.InsertCell()
        builder.Write("Row 1, Cell 1 Content.")
        ' Build the second cell
        builder.InsertCell()
        builder.Write("Row 1, Cell 2 Content.")
        ' Call the following method to end the row and start a new row.
        builder.EndRow()

        ' Build the first cell of the second row.
        builder.InsertCell()
        builder.Write("Row 2, Cell 1 Content")

        ' Build the second cell.
        builder.InsertCell()
        builder.Write("Row 2, Cell 2 Content.")
        builder.EndRow()

        ' Signal that we have finished building the table.
        builder.EndTable()

        dataDir = dataDir & Convert.ToString("DocumentBuilder.CreateSimpleTable_out_.doc")
        ' Save the document to disk.
        doc.Save(dataDir)
        ' ExEnd:SimpleTable
        Console.WriteLine(Convert.ToString(vbLf & "Simple table created successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
    Private Shared Sub FormattedTable(dataDir As String)
        ' ExStart:FormattedTable
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        Dim table As Table = builder.StartTable()
        ' Make the header row.
        builder.InsertCell()

        ' Set the left indent for the table. Table wide formatting must be applied after 
        ' at least one row is present in the table.
        table.LeftIndent = 20.0

        ' Set height and define the height rule for the header row.
        builder.RowFormat.Height = 40.0
        builder.RowFormat.HeightRule = HeightRule.AtLeast

        ' Some special features for the header row.
        builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241)
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center
        builder.Font.Size = 16
        builder.Font.Name = "Arial"
        builder.Font.Bold = True

        builder.CellFormat.Width = 100.0
        builder.Write("Header Row," & vbLf & " Cell 1")

        ' We don't need to specify the width of this cell because it's inherited from the previous cell.
        builder.InsertCell()
        builder.Write("Header Row," & vbLf & " Cell 2")

        builder.InsertCell()
        builder.CellFormat.Width = 200.0
        builder.Write("Header Row," & vbLf & " Cell 3")
        builder.EndRow()

        ' Set features for the other rows and cells.
        builder.CellFormat.Shading.BackgroundPatternColor = Color.White
        builder.CellFormat.Width = 100.0
        builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center

        ' Reset height and define a different height rule for table body
        builder.RowFormat.Height = 30.0
        builder.RowFormat.HeightRule = HeightRule.Auto
        builder.InsertCell()
        ' Reset font formatting.
        builder.Font.Size = 12
        builder.Font.Bold = False

        ' Build the other cells.
        builder.Write("Row 1, Cell 1 Content")
        builder.InsertCell()
        builder.Write("Row 1, Cell 2 Content")

        builder.InsertCell()
        builder.CellFormat.Width = 200.0
        builder.Write("Row 1, Cell 3 Content")
        builder.EndRow()

        builder.InsertCell()
        builder.CellFormat.Width = 100.0
        builder.Write("Row 2, Cell 1 Content")

        builder.InsertCell()
        builder.Write("Row 2, Cell 2 Content")

        builder.InsertCell()
        builder.CellFormat.Width = 200.0
        builder.Write("Row 2, Cell 3 Content.")
        builder.EndRow()
        builder.EndTable()

        dataDir = dataDir & Convert.ToString("DocumentBuilder.CreateFormattedTable_out_.doc")
        ' Save the document to disk.
        doc.Save(dataDir)
        ' ExEnd:FormattedTable
        Console.WriteLine(Convert.ToString(vbLf & "Formatted table created successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
    Private Shared Sub NestedTable(dataDir As String)
        ' ExStart:NestedTable
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' Build the outer table.
        Dim cell As Cell = builder.InsertCell()
        builder.Writeln("Outer Table Cell 1")

        builder.InsertCell()
        builder.Writeln("Outer Table Cell 2")

        ' This call is important in order to create a nested table within the first table
        ' Without this call the cells inserted below will be appended to the outer table.
        builder.EndTable()

        ' Move to the first cell of the outer table.
        builder.MoveTo(cell.FirstParagraph)

        ' Build the inner table.
        builder.InsertCell()
        builder.Writeln("Inner Table Cell 1")
        builder.InsertCell()
        builder.Writeln("Inner Table Cell 2")
        builder.EndTable()

        dataDir = dataDir & Convert.ToString("DocumentBuilder.InsertNestedTable_out_.doc")
        ' Save the document to disk.
        doc.Save(dataDir)
        ' ExEnd:NestedTable
        Console.WriteLine(Convert.ToString(vbLf & "Nested table created successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
