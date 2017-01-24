Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports System.Reflection

Imports Aspose.Words

Public Class MultipleDocsInMailMerge
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()

        ' Open the database connection.
        Dim connString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dataDir & "Customers.mdb"
        Dim conn As New OleDbConnection(connString)

        Try
            conn.Open()
            ' Get data from a database.
            Dim cmd As New OleDbCommand("SELECT * FROM Customers", conn)
            Dim da As New OleDbDataAdapter(cmd)
            Dim data As New DataTable()
            da.Fill(data)

            ' Open the template document.
            Dim doc As New Document(dataDir & "TestFile.Multiple Pages.doc")

            Dim counter As Integer = 1
            ' Loop though all records in the data source.
            For Each row As DataRow In data.Rows
                ' Clone the template instead of loading it from disk (for speed).
                Dim dstDoc As Document = CType(doc.Clone(True), Document)

                ' Execute mail merge.
                dstDoc.MailMerge.Execute(row)

                ' Save the document.
                dstDoc.Save(String.Format(dataDir & "TestFile.Multiple Pages_out {0}.doc", counter))
                counter += 1
            Next row
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        Finally
            ' Close the database.
            conn.Close()
        End Try

        Console.WriteLine(vbNewLine + "Mail merge performed and created multiple pages successfully." + vbNewLine + "File saved at " + dataDir + "TestFile.Multiple Pages_out.doc")
    End Sub
End Class
