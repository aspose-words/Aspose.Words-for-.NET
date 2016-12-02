Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Drawing
Imports System.Diagnostics
Imports Aspose.Words
Imports Aspose.Words.Tables
Public Class InsertTableFromHtml
    Public Shared Sub Run()
        ' ExStart:InsertTableFromHtml
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithTables()
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' Insert the table from HTML. Note that AutoFitSettings does not apply to tables
        ' Inserted from HTML.
        builder.InsertHtml("<table>" + "<tr>" + "<td>Row 1, Cell 1</td>" + "<td>Row 1, Cell 2</td>" + "</tr>" + "<tr>" + "<td>Row 2, Cell 2</td>" + "<td>Row 2, Cell 2</td>" + "</tr>" + "</table>")

        dataDir = dataDir & Convert.ToString("DocumentBuilder.InsertTableFromHtml_out.doc")
        ' Save the document to disk.
        doc.Save(dataDir)
        ' ExEnd:InsertTableFromHtml

        Console.WriteLine(Convert.ToString(vbLf & "Table inserted successfully from html." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
