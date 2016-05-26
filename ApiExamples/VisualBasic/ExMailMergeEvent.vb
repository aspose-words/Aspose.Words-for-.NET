' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.IO

Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.MailMerging

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExMailMergeEvent
		Inherits ApiExampleBase
		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub MailMergeInsertHtmlCaller()
			Me.MailMergeInsertHtml()
		End Sub

		'ExStart
		'ExFor:DocumentBuilder.InsertHtml(string)
		'ExFor:MailMerge.FieldMergingCallback
		'ExFor:IFieldMergingCallback
		'ExFor:FieldMergingArgs
		'ExFor:FieldMergingArgsBase.DocumentFieldName
		'ExFor:FieldMergingArgsBase.Document
		'ExFor:FieldMergingArgsBase.FieldValue
		'ExFor:IFieldMergingCallback.FieldMerging
		'ExFor:FieldMergingArgs.Text
		'ExSummary:Shows how to mail merge HTML data into a document.
		' File 'MailMerge.InsertHtml.doc' has merge field named 'htmlField1' in it.
		' File 'MailMerge.HtmlData.html' contains some valid Html data.
		' The same approach can be used when merging HTML data from database.
		Public Sub MailMergeInsertHtml()
			Dim doc As New Document(MyDir & "MailMerge.InsertHtml.doc")

			' Add a handler for the MergeField event.
			doc.MailMerge.FieldMergingCallback = New HandleMergeFieldInsertHtml()

			' Load some Html from file.
			Dim sr As StreamReader = File.OpenText(MyDir & "MailMerge.HtmlData.html")
			Dim htmltext As String = sr.ReadToEnd()
			sr.Close()

			' Execute mail merge.
			doc.MailMerge.Execute(New String() { "htmlField1" }, New String() { htmltext })

			' Save resulting document with a new name.
			doc.Save(MyDir & "\Artifacts\MailMerge.InsertHtml.doc")
		End Sub

		Private Class HandleMergeFieldInsertHtml
			Implements IFieldMergingCallback
			''' <summary>
			''' This is called when merge field is actually merged with data in the document.
			''' </summary>
			Private Sub IFieldMergingCallback_FieldMerging(ByVal e As FieldMergingArgs) Implements IFieldMergingCallback.FieldMerging
				' All merge fields that expect HTML data should be marked with some prefix, e.g. 'html'.
				If e.DocumentFieldName.StartsWith("html") Then
					' Insert the text for this merge field as HTML data, using DocumentBuilder.
					Dim builder As New DocumentBuilder(e.Document)
					builder.MoveToMergeField(e.DocumentFieldName)
					builder.InsertHtml(CStr(e.FieldValue))

					' The HTML text itself should not be inserted.
					' We have already inserted it as an HTML.
					e.Text = ""
				End If
			End Sub

			Private Sub ImageFieldMerging(ByVal e As ImageFieldMergingArgs) Implements IFieldMergingCallback.ImageFieldMerging
				' Do nothing.
			End Sub
		End Class
		'ExEnd


		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub MailMergeInsertCheckBoxCaller()
			Me.MailMergeInsertCheckBox()
		End Sub

		'ExStart
		'ExFor:DocumentBuilder.MoveToMergeField(string)
		'ExFor:DocumentBuilder.InsertCheckBox
		'ExFor:FieldMergingArgsBase.FieldName
		'ExSummary:Shows how to insert checkbox form fields into a document during mail merge.
		' File 'MailMerge.InsertCheckBox.doc' is a template
		' containing the table with the following fields in it:
		' <<TableStart:StudentCourse>> <<CourseName>> <<TableEnd:StudentCourse>>.
		Public Sub MailMergeInsertCheckBox()
			Dim doc As New Document(MyDir & "MailMerge.InsertCheckBox.doc")

			' Add a handler for the MergeField event.
			doc.MailMerge.FieldMergingCallback = New HandleMergeFieldInsertCheckBox()

			' Execute mail merge with regions.
			Dim dataTable As DataTable = GetStudentCourseDataTable()
			doc.MailMerge.ExecuteWithRegions(dataTable)

			' Save resulting document with a new name.
			doc.Save(MyDir & "\Artifacts\MailMerge.InsertCheckBox.doc")
		End Sub

		Private Class HandleMergeFieldInsertCheckBox
			Implements IFieldMergingCallback
			''' <summary>
			''' This is called for each merge field in the document
			''' when Document.MailMerge.ExecuteWithRegions is called.
			''' </summary>
			Private Sub IFieldMergingCallback_FieldMerging(ByVal e As FieldMergingArgs) Implements IFieldMergingCallback.FieldMerging
				If e.DocumentFieldName.Equals("CourseName") Then
					' Insert the checkbox for this merge field, using DocumentBuilder.
					Dim builder As New DocumentBuilder(e.Document)
					builder.MoveToMergeField(e.FieldName)
					builder.InsertCheckBox(e.DocumentFieldName + Me.mCheckBoxCount.ToString(), False, 0)
					builder.Write(CStr(e.FieldValue))
					Me.mCheckBoxCount += 1
				End If
			End Sub

			Private Sub ImageFieldMerging(ByVal args As ImageFieldMergingArgs) Implements IFieldMergingCallback.ImageFieldMerging
				' Do nothing.
			End Sub

			''' <summary>
			''' Counter for CheckBox name generation
			''' </summary>
			Private mCheckBoxCount As Integer
		End Class

		''' <summary>
		''' Create DataTable and fill it with data.
		''' In real life this DataTable should be filled from a database.
		''' </summary>
		Private Shared Function GetStudentCourseDataTable() As DataTable
			Dim dataTable As New DataTable("StudentCourse")
			dataTable.Columns.Add("CourseName")
			For i As Integer = 0 To 9
				Dim datarow As DataRow = dataTable.NewRow()
				dataTable.Rows.Add(datarow)
				datarow(0) = "Course " & i.ToString()
			Next i
			Return dataTable
		End Function
		'ExEnd

		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub MailMergeAlternatingRowsCaller()
			Me.MailMergeAlternatingRows()
		End Sub

		'ExStart
		'ExId:MailMergeAlternatingRows
		'ExSummary:Demonstrates how to implement custom logic in the MergeField event to apply cell formatting.
		Public Sub MailMergeAlternatingRows()
			Dim doc As New Document(MyDir & "MailMerge.AlternatingRows.doc")

			' Add a handler for the MergeField event.
			doc.MailMerge.FieldMergingCallback = New HandleMergeFieldAlternatingRows()

			' Execute mail merge with regions.
			Dim dataTable As DataTable = GetSuppliersDataTable()
			doc.MailMerge.ExecuteWithRegions(dataTable)

			doc.Save(MyDir & "\Artifacts\MailMerge.AlternatingRows.doc")
		End Sub

		Private Class HandleMergeFieldAlternatingRows
			Implements IFieldMergingCallback
			''' <summary>
			''' Called for every merge field encountered in the document.
			''' We can either return some data to the mail merge engine or do something
			''' else with the document. In this case we modify cell formatting.
			''' </summary>
			Private Sub IFieldMergingCallback_FieldMerging(ByVal e As FieldMergingArgs) Implements IFieldMergingCallback.FieldMerging
				If Me.mBuilder Is Nothing Then
					Me.mBuilder = New DocumentBuilder(e.Document)
				End If

				' This way we catch the beginning of a new row.
				If e.FieldName.Equals("CompanyName") Then
					' Select the color depending on whether the row number is even or odd.
					Dim rowColor As Color
					If IsOdd(Me.mRowIdx) Then
						rowColor = Color.FromArgb(213, 227, 235)
					Else
						rowColor = Color.FromArgb(242, 242, 242)
					End If

					' There is no way to set cell properties for the whole row at the moment,
					' so we have to iterate over all cells in the row.
					For colIdx As Integer = 0 To 3
						Me.mBuilder.MoveToCell(0, Me.mRowIdx, colIdx, 0)
						Me.mBuilder.CellFormat.Shading.BackgroundPatternColor = rowColor
					Next colIdx

					Me.mRowIdx += 1
				End If
			End Sub

			Private Sub ImageFieldMerging(ByVal args As ImageFieldMergingArgs) Implements IFieldMergingCallback.ImageFieldMerging
				' Do nothing.
			End Sub

			Private mBuilder As DocumentBuilder
			Private mRowIdx As Integer
		End Class

		''' <summary>
		''' Returns true if the value is odd; false if the value is even.
		''' </summary>
		Private Shared Function IsOdd(ByVal value As Integer) As Boolean
			' The code is a bit complex, but otherwise automatic conversion to VB does not work.
			Return ((value \ 2) * 2).Equals(value)
		End Function

		''' <summary>
		''' Create DataTable and fill it with data.
		''' In real life this DataTable should be filled from a database.
		''' </summary>
		Private Shared Function GetSuppliersDataTable() As DataTable
			Dim dataTable As New DataTable("Suppliers")
			dataTable.Columns.Add("CompanyName")
			dataTable.Columns.Add("ContactName")
			For i As Integer = 0 To 9
				Dim datarow As DataRow = dataTable.NewRow()
				dataTable.Rows.Add(datarow)
				datarow(0) = "Company " & i.ToString()
				datarow(1) = "Contact " & i.ToString()
			Next i
			Return dataTable
		End Function
		'ExEnd

		<Test> _
		Public Sub MailMergeImageFromUrl()
			'ExStart
			'ExFor:MailMerge.Execute(String[], Object[])
			'ExSummary:Demonstrates how to merge an image from a web address using an Image field.
			Dim doc As New Document(MyDir & "MailMerge.MergeImageSimple.doc")

			' Pass a URL which points to the image to merge into the document.
			doc.MailMerge.Execute(New String() { "Logo" }, New Object() { "http://www.aspose.com/images/aspose-logo.gif" })

			doc.Save(MyDir & "\Artifacts\MailMerge.MergeImageFromUrl.doc")
			'ExEnd

			' Verify the image was merged into the document.
			Dim logoImage As Shape = CType(doc.GetChild(NodeType.Shape, 0, True), Shape)
			Assert.IsNotNull(logoImage)
			Assert.IsTrue(logoImage.HasImage)
		End Sub

		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub MailMergeImageFromBlobCaller()
			Me.MailMergeImageFromBlob()
		End Sub

		'ExStart
		'ExFor:MailMerge.FieldMergingCallback
		'ExFor:MailMerge.ExecuteWithRegions(IDataReader,string)
		'ExFor:IFieldMergingCallback
		'ExFor:ImageFieldMergingArgs
		'ExFor:IFieldMergingCallback.FieldMerging
		'ExFor:IFieldMergingCallback.ImageFieldMerging
		'ExFor:ImageFieldMergingArgs.ImageStream
		'ExId:MailMergeImageFromBlob
		'ExSummary:Shows how to insert images stored in a database BLOB field into a report.
		Public Sub MailMergeImageFromBlob()
			Dim doc As New Document(MyDir & "MailMerge.MergeImage.doc")

			' Set up the event handler for image fields.
			doc.MailMerge.FieldMergingCallback = New HandleMergeImageFieldFromBlob()

			' Open a database connection.
			Dim connString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DatabaseDir & "Northwind.mdb"
			Dim conn As New OleDbConnection(connString)
			conn.Open()

			' Open the data reader. It needs to be in the normal mode that reads all record at once.
			Dim cmd As New OleDbCommand("SELECT * FROM Employees", conn)
			Dim dataReader As IDataReader = cmd.ExecuteReader()

			' Perform mail merge.
			doc.MailMerge.ExecuteWithRegions(dataReader, "Employees")

			' Close the database.
			conn.Close()

			doc.Save(MyDir & "\Artifacts\MailMerge.MergeImage.doc")
		End Sub

		Private Class HandleMergeImageFieldFromBlob
			Implements IFieldMergingCallback
			Private Sub IFieldMergingCallback_FieldMerging(ByVal args As FieldMergingArgs) Implements IFieldMergingCallback.FieldMerging
				' Do nothing.
			End Sub

			''' <summary>
			''' This is called when mail merge engine encounters Image:XXX merge field in the document.
			''' You have a chance to return an Image object, file name or a stream that contains the image.
			''' </summary>
			Private Sub ImageFieldMerging(ByVal e As ImageFieldMergingArgs) Implements IFieldMergingCallback.ImageFieldMerging
				' The field value is a byte array, just cast it and create a stream on it.
				Dim imageStream As New MemoryStream(CType(e.FieldValue, Byte()))
				' Now the mail merge engine will retrieve the image from the stream.
				e.ImageStream = imageStream
			End Sub
		End Class
		'ExEnd
	End Class
End Namespace
