' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.Web
Imports System.Collections

Imports Aspose.Words.Fields
Imports Aspose.Words
Imports Aspose.Words.MailMerging

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExMailMerge
		Inherits ApiExampleBase
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
			doc.Save(Response, "\Artifacts\MailMerge.ExecuteArray.doc", ContentDisposition.Inline, Nothing)
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

			doc.Save(MyDir & "\Artifacts\MailMerge.ExecuteDataTable.doc")
			'ExEnd
		End Sub

		<Test, TestCase(True, "first line" & Constants.vbCr & "second line" & Constants.vbCr & "third line" & Constants.vbFormFeed), TestCase(False, " first line" & Constants.vbCr & "second line" & Constants.vbCr & "third line " & Constants.vbFormFeed)> _
		Public Sub TrimWhiteSpaces(ByVal [option] As Boolean, ByVal expectedText As String)
			Dim doc As New Document()
			Dim builder As New DocumentBuilder(doc)

			builder.InsertField("MERGEFIELD field", Nothing)

			doc.MailMerge.TrimWhitespaces = [option]
			doc.MailMerge.Execute(New String() { "field" }, New Object() { " first line" & Constants.vbCr & "second line" & Constants.vbCr & "third line " })

			Assert.AreEqual(expectedText, doc.GetText())
		End Sub

		<Test> _
		Public Sub ExecuteDataReader()
			'ExStart
			'ExFor:MailMerge.Execute(IDataReader)
			'ExSummary:Executes mail merge from an ADO.NET DataReader.
			' Open the template document
			Dim doc As New Document(MyDir & "MailingLabelsDemo.doc")

			' Open the database connection.
			Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DatabaseDir & "Northwind.mdb"
			Dim conn As New OleDbConnection(connString)
			Try
				conn.Open()
			Catch ex As Exception
				Debug.WriteLine(ex)
			End Try


			' Open the data reader.
			Dim cmd As New OleDbCommand("SELECT TOP 50 * FROM Customers ORDER BY Country, CompanyName", conn)
			Dim dataReader As OleDbDataReader = cmd.ExecuteReader()

			' Perform the mail merge
			doc.MailMerge.Execute(dataReader)

			' Close database.
			dataReader.Close()
			conn.Close()

			doc.Save(MyDir & "\Artifacts\MailMerge.ExecuteDataReader.doc")
			'ExEnd
		End Sub

		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub ExecuteDataViewCaller()
			Me.ExecuteDataView()
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

			doc.Save(MyDir & "\Artifacts\MailMerge.ExecuteDataView.doc")
		End Sub

		Private Shared Function GetOrders() As DataTable
			' Open a database connection.
			Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DatabaseDir & "Northwind.mdb"
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

			doc.Save(MyDir & "\Artifacts\MailMerge.ExecuteWithRegionsDataSet.doc")
			'ExEnd
		End Sub


		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub ExecuteWithRegionsDataTableCaller()
			Me.ExecuteWithRegionsDataTable()
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

			doc.Save(MyDir & "\Artifacts\MailMerge.ExecuteWithRegionsDataTable.doc")
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
			Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DatabaseDir & "Northwind.mdb"
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

        <Test, TestCase(True, "{{ testfield1 }}value 1{{ testfield3 }}" & Constants.vbFormFeed), TestCase(False, ChrW(&H13) & "MERGEFIELD ""testfield1""" & ChrW(&H14) & "«testfield1»" & ChrW(&H15) & "value 1" & ChrW(&H13) & "MERGEFIELD ""testfield3""" & ChrW(&H14) & "«testfield3»" & ChrW(&H15) & Constants.vbFormFeed)> _
        Public Sub MustasheTemplateSyntax(ByVal restoreTags As Boolean, ByVal sectionText As String)
            Dim doc As New Document()
            Dim builder As New DocumentBuilder(doc)
            builder.Write("{{ testfield1 }}")
            builder.Write("{{ testfield2 }}")
            builder.Write("{{ testfield3 }}")

            doc.MailMerge.UseNonMergeFields = True
            doc.MailMerge.PreserveUnusedTags = restoreTags

            Dim table As New DataTable("Test")
            table.Columns.Add("testfield2")
            table.Rows.Add(New Object() {"value 1"})

            doc.MailMerge.Execute(table)

            Dim paraText As String = DocumentHelper.GetParagraphText(doc, 0)

            Assert.AreEqual(sectionText, paraText)
        End Sub

		<Test> _
		Public Sub TestMailMergeGetRegionsHierarchy()
			'ExStart
			'ExFor:MailMerge.GetRegionsHierarchy
			'ExFor:MailMergeRegionInfo.Regions
			'ExFor:MailMergeRegionInfo.Name
			'ExFor:MailMergeRegionInfo.Fields
			'ExFor:MailMergeRegionInfo.StartField
			'ExFor:MailMergeRegionInfo.EndField
			'ExSummary:Shows how to get MailMergeRegionInfo and work with it
			Dim doc As New Document(MyDir & "MailMerge.TestRegionsHierarchy.doc")

			'Returns a full hierarchy of regions (with fields) available in the document.
			Dim regionInfo As MailMergeRegionInfo = doc.MailMerge.GetRegionsHierarchy()

			'Get top regions in the document
			Dim topRegions As ArrayList = regionInfo.Regions
			Assert.AreEqual(2, topRegions.Count)
			Assert.AreEqual((CType(topRegions(0), MailMergeRegionInfo)).Name, "Region1")
			Assert.AreEqual((CType(topRegions(1), MailMergeRegionInfo)).Name, "Region2")

			'Get nested region in first top region
			Dim nestedRegions As ArrayList = (CType(topRegions(0), MailMergeRegionInfo)).Regions
			Assert.AreEqual(2, nestedRegions.Count)
			Assert.AreEqual((CType(nestedRegions(0), MailMergeRegionInfo)).Name, "NestedRegion1")
			Assert.AreEqual((CType(nestedRegions(1), MailMergeRegionInfo)).Name, "NestedRegion2")

			'Get field list in first top region
			Dim fieldList As ArrayList = (CType(topRegions(0), MailMergeRegionInfo)).Fields
			Assert.AreEqual(4, fieldList.Count)

			Dim startFieldMergeField As FieldMergeField = (CType(nestedRegions(0), MailMergeRegionInfo)).StartField
			Assert.AreEqual("TableStart:NestedRegion1", startFieldMergeField.FieldName)

			Dim endFieldMergeField As FieldMergeField = (CType(nestedRegions(0), MailMergeRegionInfo)).EndField
			Assert.AreEqual("TableEnd:NestedRegion1", endFieldMergeField.FieldName)
			'ExEnd
		End Sub

		<Test> _
		Public Sub TestTagsReplacedEventShouldRisedWithUseNonMergeFieldsOption()
			Dim document As New Document()
			document.MailMerge.UseNonMergeFields = True

			Dim mailMergeCallbackStub As New MailMergeCallbackStub()
			document.MailMerge.MailMergeCallback = mailMergeCallbackStub

			document.MailMerge.Execute(New String(){}, New Object(){})

			Assert.AreEqual(1, mailMergeCallbackStub.TagsReplacedCounter)
		End Sub

		Private Class MailMergeCallbackStub
			Implements IMailMergeCallback
            Public Sub TagsReplaced() Implements IMailMergeCallback.TagsReplaced
                mTagsReplacedCounter += 1
            End Sub

			Public ReadOnly Property TagsReplacedCounter() As Integer
				Get
					Return mTagsReplacedCounter
				End Get
			End Property

			Private mTagsReplacedCounter As Integer
        End Class
	End Class
End Namespace
