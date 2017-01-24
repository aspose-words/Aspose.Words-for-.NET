Imports Aspose.Words
Imports Aspose.Words.MailMerging
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Web
Public Class ProduceMultipleDocuments
    Public Shared Sub Run()
        ' ExStart:ProduceMultipleDocuments            
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()
        ' Open the database connection.
        Dim connString As String = (Convert.ToString("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=") & dataDir) + "Customers.mdb"
        Dim conn As New OleDbConnection(connString)
        conn.Open()
        ' Get data from a database.
        Dim cmd As New OleDbCommand("SELECT * FROM Customers", conn)
        Dim da As New OleDbDataAdapter(cmd)
        Dim data As New DataTable()
        da.Fill(data)

        ' Open the template document.
        Dim doc As New Document(dataDir & Convert.ToString("TestFile.doc"))

        Dim counter As Integer = 1
        ' Loop though all records in the data source.
        For Each row As DataRow In data.Rows
            ' Clone the template instead of loading it from disk (for speed).
            Dim dstDoc As Document = DirectCast(doc.Clone(True), Document)

            ' Execute mail merge.
            dstDoc.MailMerge.Execute(row)

            ' Save the document.
            dstDoc.Save(String.Format(dataDir & Convert.ToString("TestFile_out{0}.doc"), System.Math.Max(System.Threading.Interlocked.Increment(counter), counter - 1)))
        Next
        ' ExEnd:ProduceMultipleDocuments
        Console.WriteLine(Convert.ToString(vbLf & "Produce multiple documents performed successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
