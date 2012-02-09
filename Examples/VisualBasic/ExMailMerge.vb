'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Web
Imports Aspose.Words
'ExStart
'ExId:UsingReportingNamespace
'ExSummary:Include the following statement in your code if you are using mail merge functionality.
Imports Aspose.Words.Reporting
'ExEnd

Imports NUnit.Framework

Namespace Examples
	<TestFixture> _
	Public Class ExMailMerge
		Inherits ExBase
		<Test, ExpectedException(GetType(ArgumentNullException))> _
		Public Sub ExecuteArray()
			Dim Response As HttpResponse = Nothing

			'ExStart
			'ExFor:MailMerge.Execute(String[],Object[])
			'ExFor:ContentDisposition
			'ExFor:Document.Save(HttpResponse,String,ContentDisposition,SaveOptions)
			'ExId:MailMergeArray
			'ExSummary:Performs a simple insertion of data into merge fields and sends the document to the browser inline.
			' Open an existing document.
			Dim doc As New Document(MyDir & "MailMerge.ExecuteArray.doc")

			' Fill the fields in the document with user data.
			doc.MailMerge.Execute(New String() {"FullName", "Company", "Address", "Address2", "City"}, New Object() {"James Bond", "MI5 Headquarters", "Milbank", "", "London"})

			' Send the document in Word format to the client browser with an option to save to disk or open inside the current browser.
			doc.Save(Response, "MailMerge.ExecuteArray Out.doc", ContentDisposition.Inline, Nothing)
			'ExEnd
		End Sub

		<Test> _
		Public Sub ExecuteDataTable()
			'ExStart
			'ExFor:Document
			'ExFor:MailMerge
			'ExFor:MailMerge.Execute(DataTable)
			'ExFor:Document.MailMerge
			'ExSummary:Executes mail merge from an ADO.NET DataTable.
			Dim doc As New Document(MyDir & "MailMerge.ExecuteDataTable.doc")

			' This example creates a table, but you would normally load table from a database. 
			Dim table As New DataTable("Test")
			table.Columns.Add("CustomerName")
			table.Columns.Add("Address")
			table.Rows.Add(New Object() {"Thomas Hardy", "120 Hanover Sq., London"})
			table.Rows.Add(New Object() {"Paolo Accorti", "Via Monte Bianco 34, Torino"})

			' Field values from the table are inserted into the mail merge fields found in the document.
			doc.MailMerge.Execute(table)

			doc.Save(MyDir & "MailMerge.ExecuteDataTable Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub ExecuteDataReader()
			'ExStart
			'ExFor:MailMerge.Execute(IDataReader)
			'ExSummary:Executes mail merge from an ADO.NET DataReader.
			' Open the template document
			Dim doc As New Document(MyDir & "MailingLabelsDemo.doc")

			' Open the database connection.
			Dim connString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabaseDir & "Northwind.mdb"
			Dim conn As New OleDbConnection(connString)
			conn.Open()

			' Open the data reader.
			Dim cmd As New OleDbCommand("SELECT TOP 50 * FROM Customers ORDER BY Country, CompanyName", conn)
			Dim dataReader As OleDbDataReader = cmd.ExecuteReader()

			' Perform the mail merge
			doc.MailMerge.Execute(dataReader)

			' Close database.
			dataReader.Close()
			conn.Close()

			doc.Save(MyDir & "MailMerge.ExecuteDataReader Out.doc")
			'ExEnd
		End Sub

		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub ExecuteDataViewCaller()
			ExecuteDataView()
		End Sub

		'ExStart
		'ExFor:MailMerge.Execute(DataView)
		'ExSummary:Executes mail merge from an ADO.NET DataView.
		Public Sub ExecuteDataView()
			' Open the document that we want to fill with data.
			Dim doc As New Document(MyDir & "MailMerge.ExecuteDataView.doc")

			' Get the data from the database.
			Dim orderTable As DataTable = GetOrders()

			' Create a customized view of the data.
			Dim orderView As New DataView(orderTable)
			orderView.RowFilter = "OrderId = 10444"

			' Populate the document with the data.
			doc.MailMerge.Execute(orderView)

			doc.Save(MyDir & "MailMerge.ExecuteDataView Out.doc")
		End Sub

		Private Shared Function GetOrders() As DataTable
			' Open a database connection.
			Dim connString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabaseDir & "Northwind.mdb"
			Dim conn As New OleDbConnection(connString)
			conn.Open()

			' Create the command.
			Dim cmd As New OleDbCommand("SELECT * FROM AsposeWordOrders", conn)

			' Fill an ADO.NET table from the command.
			Dim da As New OleDbDataAdapter(cmd)
			Dim table As New DataTable()
			da.Fill(table)

			' Close database.
			conn.Close()

			Return table
		End Function
		'ExEnd


		<Test> _
		Public Sub ExecuteWithRegionsDataSet()
			'ExStart
			'ExFor:MailMerge.ExecuteWithRegions(DataSet)
			'ExSummary:Executes a mail merge with repeatable regions from an ADO.NET DataSet.
			' Open the document. 
			' For a mail merge with repeatable regions, the document should have mail merge regions 
			' in the document designated with MERGEFIELD TableStart:MyTableName and TableEnd:MyTableName.
			Dim doc As New Document(MyDir & "MailMerge.ExecuteWithRegions.doc")

			Dim orderId As Integer = 10444

			' Populate tables and add them to the dataset.
			' For a mail merge with repeatable regions, DataTable.TableName should be 
			' set to match the name of the region defined in the document.
			Dim dataSet As New DataSet()

			Dim orderTable As DataTable = GetTestOrder(orderId)
			dataSet.Tables.Add(orderTable)

			Dim orderDetailsTable As DataTable = GetTestOrderDetails(orderId)
			dataSet.Tables.Add(orderDetailsTable)

			' This looks through all mail merge regions inside the document and for each
			' region tries to find a DataTable with a matching name inside the DataSet.
			' If a table is found, its content is merged into the mail merge region in the document.
			doc.MailMerge.ExecuteWithRegions(dataSet)

			doc.Save(MyDir & "MailMerge.ExecuteWithRegionsDataSet Out.doc")
			'ExEnd
		End Sub


		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub ExecuteWithRegionsDataTableCaller()
			ExecuteWithRegionsDataTable()
		End Sub

		'ExStart
		'ExFor:Document.MailMerge
		'ExFor:MailMerge.ExecuteWithRegions(DataTable)
		'ExFor:MailMerge.ExecuteWithRegions(DataView)
		'ExId:MailMergeRegions
		'ExSummary:Executes a mail merge with repeatable regions.
		Public Sub ExecuteWithRegionsDataTable()
			Dim doc As New Document(MyDir & "MailMerge.ExecuteWithRegions.doc")

			Dim orderId As Integer = 10444

			' Perform several mail merge operations populating only part of the document each time.

			' Use DataTable as a data source.
			Dim orderTable As DataTable = GetTestOrder(orderId)
			doc.MailMerge.ExecuteWithRegions(orderTable)

			' Instead of using DataTable you can create a DataView for custom sort or filter and then mail merge.
			Dim orderDetailsView As New DataView(GetTestOrderDetails(orderId))
			orderDetailsView.Sort = "ExtendedPrice DESC"
			doc.MailMerge.ExecuteWithRegions(orderDetailsView)

			doc.Save(MyDir & "MailMerge.ExecuteWithRegionsDataTable Out.doc")
		End Sub

		Private Shared Function GetTestOrder(ByVal orderId As Integer) As DataTable
			Dim table As DataTable = ExecuteDataTable(String.Format("SELECT * FROM AsposeWordOrders WHERE OrderId = {0}", orderId))
			table.TableName = "Orders"
			Return table
		End Function

		Private Shared Function GetTestOrderDetails(ByVal orderId As Integer) As DataTable
			Dim table As DataTable = ExecuteDataTable(String.Format("SELECT * FROM AsposeWordOrderDetails WHERE OrderId = {0} ORDER BY ProductID", orderId))
			table.TableName = "OrderDetails"
			Return table
		End Function

		''' <summary>
		''' Utility function that creates a connection, command, 
		''' executes the command and return the result in a DataTable.
		''' </summary>
		Private Shared Function ExecuteDataTable(ByVal commandText As String) As DataTable
			' Open the database connection.
			Dim connString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabaseDir & "Northwind.mdb"
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
		'ExEnd

		<Test> _
		Public Sub MappedDataFields()
			Dim doc As New Document()
			'ExStart
			'ExFor:MailMerge.MappedDataFields
			'ExFor:MappedDataFieldCollection
			'ExFor:MappedDataFieldCollection.Add
			'ExId:MailMergeMappedDataFields
			'ExSummary:Shows how to add a mapping when a merge field in a document and a data field in a data source have different names.
			doc.MailMerge.MappedDataFields.Add("MyFieldName_InDocument", "MyFieldName_InDataSource")
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetFieldNames()
			Dim doc As New Document()
			'ExStart
			'ExFor:MailMerge.GetFieldNames
			'ExId:MailMergeGetFieldNames
			'ExSummary:Shows how to get names of all merge fields in a document.
			Dim fieldNames() As String = doc.MailMerge.GetFieldNames()
			'ExEnd
		End Sub

		<Test> _
		Public Sub DeleteFields()
			Dim doc As New Document()
			'ExStart
			'ExFor:MailMerge.DeleteFields
			'ExId:MailMergeDeleteFields
			'ExSummary:Shows how to delete all merge fields from a document without executing mail merge.
			doc.MailMerge.DeleteFields()
			'ExEnd
		End Sub

		<Test> _
		Public Sub RemoveContainingFields()
			Dim doc As New Document()
			'ExStart
			'ExFor:MailMerge.CleanupOptions
			'ExFor:MailMergeCleanupOptions
			'ExId:MailMergeRemoveContainingFields
			'ExSummary:Shows how to instruct the mail merge engine to remove any containing fields from around a merge field during mail merge.
			doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveContainingFields
			'ExEnd
		End Sub

		<Test> _
		Public Sub RemoveUnusedFields()
			Dim doc As New Document()
			'ExStart
			'ExFor:MailMerge.CleanupOptions
			'ExFor:MailMergeCleanupOptions
			'ExId:MailMergeRemoveUnusedFields
			'ExSummary:Shows how to automatically remove unmerged merge fields during mail merge.
			doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedFields
			'ExEnd
		End Sub

		<Test> _
		Public Sub RemoveEmptyParagraphs()
			Dim doc As New Document()
			'ExStart
			'ExFor:MailMerge.CleanupOptions
			'ExFor:MailMergeCleanupOptions
			'ExId:MailMergeRemoveEmptyParagraphs
			'ExSummary:Shows how to make sure empty paragraphs that result from merging fields with no data are removed from the document.
			doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveEmptyParagraphs
			'ExEnd
		End Sub

		<Test> _
		Public Sub UseNonMergeFields()
			Dim doc As New Document()
			'ExStart
			'ExFor:MailMerge.UseNonMergeFields
			'ExSummary:Shows how to perform mail merge into merge fields and into additional fields types.
			doc.MailMerge.UseNonMergeFields = True
			'ExEnd
		End Sub
	End Class
End Namespace
