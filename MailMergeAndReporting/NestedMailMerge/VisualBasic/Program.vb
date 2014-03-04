' Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////
'ExStart
'ExId:NestedMailMerge
'ExSummary:Shows how to generate an invoice using nested mail merge regions.

Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.IO
Imports System.Reflection
Imports System.Diagnostics

Imports Aspose.Words
Imports System.Collections

Namespace NestedMailMerge
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create the Dataset and read the XML.
			Dim pizzaDs As New DataSet()

			' Note: The Datatable.TableNames and the DataSet.Relations are defined implicitly by .NET through ReadXml.
			' To see examples of how to set up relations manually check the corresponding documentation of this sample
			pizzaDs.ReadXml(dataDir & "CustomerData.xml")

			' Open the template document.
			Dim doc As New Document(dataDir & "Invoice Template.doc")

			' Execute the nested mail merge with regions
			doc.MailMerge.ExecuteWithRegions(pizzaDs)

			' Save the output to file
			doc.Save(dataDir & "Invoice Out.doc")

			Debug.Assert(doc.MailMerge.GetFieldNames().Length = 0, "There was a problem with mail merge") 'ExSkip
		End Sub
	End Class
End Namespace
'ExEnd

Public Class DataRelationExample
	Public Shared Sub CreateRelationship()
		Dim dataSet As New DataSet()
		Dim orderTable As New DataTable()
		Dim itemTable As New DataTable()
		'ExStart
		'ExId:NestedMailMergeCreateRelationship
		'ExSummary:Shows how to create a simple DataRelation for use in nested mail merge.
		dataSet.Relations.Add(New DataRelation("OrderToItem", orderTable.Columns("Order_Id"), itemTable.Columns("Order_Id")))
		'ExEnd
	End Sub

	Public Shared Sub DisableForeignKeyConstraints()
		Dim dataSet As New DataSet()
		Dim orderTable As New DataTable()
		Dim itemTable As New DataTable()
		'ExStart
		'ExId:NestedMailMergeDisableConstraints
		'ExSummary:Shows how to disable foreign key constraints when creating a DataRelation for use in nested mail merge.
		dataSet.Relations.Add(New DataRelation("OrderToItem", orderTable.Columns("Order_Id"), itemTable.Columns("Order_Id"), False))
		'ExEnd
	End Sub
End Class