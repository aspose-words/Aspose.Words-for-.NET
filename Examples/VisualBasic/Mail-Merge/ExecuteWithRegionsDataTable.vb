Imports Aspose.Words
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Text

Class ExecuteWithRegionsDataTable
    Public Shared Sub Run()
        ' ExStart:ExecuteWithRegionsDataTable
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()
        Dim fileName As String = "MailMerge.ExecuteWithRegions.doc"
        Dim doc As New Document(dataDir & fileName)

        Dim orderId As Integer = 10444

        ' Perform several mail merge operations populating only part of the document each time.

        ' Use DataTable as a data source.
        Dim orderTable As DataTable = GetTestOrder(orderId)
        doc.MailMerge.ExecuteWithRegions(orderTable)

        ' Instead of using DataTable you can create a DataView for custom sort or filter and then mail merge.
        Dim orderDetailsView As New DataView(GetTestOrderDetails(orderId))
        orderDetailsView.Sort = "ExtendedPrice DESC"
        doc.MailMerge.ExecuteWithRegions(orderDetailsView)

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        doc.Save(dataDir)
        ' ExEnd:ExecuteWithRegionsDataTable
        Console.WriteLine(Convert.ToString(vbLf & "Mail merge executed successfully with repeatable regions." & vbLf & "File saved at ") & dataDir)
    End Sub
    ' ExStart:ExecuteWithRegionsDataTableMethods
    Private Shared Function GetTestOrder(orderId As Integer) As DataTable
        Dim table As DataTable = ExecuteDataTable(String.Format("SELECT * FROM AsposeWordOrders WHERE OrderId = {0}", orderId))
        table.TableName = "Orders"
        Return table
    End Function
    Private Shared Function GetTestOrderDetails(orderId As Integer) As DataTable
        Dim table As DataTable = ExecuteDataTable(String.Format("SELECT * FROM AsposeWordOrderDetails WHERE OrderId = {0} ORDER BY ProductID", orderId))
        table.TableName = "OrderDetails"
        Return table
    End Function
    ''' <summary>
    ''' Utility function that creates a connection, command, 
    ''' Executes the command and return the result in a DataTable.
    ''' </summary>
    Private Shared Function ExecuteDataTable(commandText As String) As DataTable
        ' Open the database connection.
        Dim connString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + RunExamples.GetDataDir_Database() + "Northwind.mdb"
        Dim conn As New OleDbConnection(connString)
        conn.Open()

        ' Create and execute a command.
        Dim cmd As New OleDbCommand(commandText, conn)
        Dim da As New OleDbDataAdapter(cmd)
        Dim table As New DataTable()
        da.Fill(table)

        ' Close the database.
        conn.Close()

        Return table
    End Function
    ' ExEnd:ExecuteWithRegionsDataTableMethods
End Class
