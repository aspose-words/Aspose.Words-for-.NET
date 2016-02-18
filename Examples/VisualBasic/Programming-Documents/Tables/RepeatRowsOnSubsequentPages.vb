Imports System.Collections
Imports System.IO
Imports System.Drawing
Imports Aspose.Words
Imports Aspose.Words.Tables
Public Class RepeatRowsOnSubsequentPages
    Public Shared Sub Run()
        ' ExStart:RepeatRowsOnSubsequentPages
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithTables()

        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        Dim table As Table = builder.StartTable()
        builder.RowFormat.HeadingFormat = True
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center
        builder.CellFormat.Width = 100
        builder.InsertCell()
        builder.Writeln("Heading row 1")
        builder.EndRow()
        builder.InsertCell()
        builder.Writeln("Heading row 2")
        builder.EndRow()

        builder.CellFormat.Width = 50
        builder.ParagraphFormat.ClearFormatting()

        ' Insert some content so the table is long enough to continue onto the next page.
        For i As Integer = 0 To 49
            builder.InsertCell()
            builder.RowFormat.HeadingFormat = False
            builder.Write("Column 1 Text")
            builder.InsertCell()
            builder.Write("Column 2 Text")
            builder.EndRow()
        Next

        dataDir = dataDir & Convert.ToString("Table.HeadingRow_out_.doc")
        ' Save the document to disk.
        doc.Save(dataDir)
        ' ExEnd:RepeatRowsOnSubsequentPages
        Console.WriteLine(Convert.ToString(vbLf & "Table build successfully which include heading rows that repeat on subsequent pages.." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
