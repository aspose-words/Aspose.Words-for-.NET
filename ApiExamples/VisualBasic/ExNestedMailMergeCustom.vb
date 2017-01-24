' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System.Collections

Imports Aspose.Words
Imports Aspose.Words.MailMerging

Imports NUnit.Framework

Namespace ApiExamples
	<TestFixture> _
	Public Class ExNestedMailMergeCustom
		Inherits ApiExampleBase
		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub MailMergeCustomDataSourceCaller()
			Me.MailMergeCustomDataSource()
		End Sub

		Public Sub MailMergeCustomDataSource()
			' Create some data that we will use in the mail merge.
			Dim customers As New CustomerList()
			customers.Add(New Customer("Thomas Hardy", "120 Hanover Sq., London"))
			customers.Add(New Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"))

			' Create some data for nesting in the mail merge.
			customers(0).Orders.Add(New Order("Rugby World Cup Cap", 2))
			customers(0).Orders.Add(New Order("Rugby World Cup Ball", 1))
			customers(1).Orders.Add(New Order("Rugby World Cup Guide", 1))

			' Open the template document.
			Dim doc As New Document(MyDir & "NestedMailMerge.CustomDataSource.doc")

			' To be able to mail merge from your own data source, it must be wrapped
			' into an object that implements the IMailMergeDataSource interface.
			Dim customersDataSource As New CustomerMailMergeDataSource(customers)

			' Now you can pass your data source into Aspose.Words.
			doc.MailMerge.ExecuteWithRegions(customersDataSource)

			doc.Save(MyDir & "\Artifacts\NestedMailMerge.CustomDataSource.doc")
		End Sub

		''' <summary>
		''' An example of a "data entity" class in your application.
		''' </summary>
		Public Class Customer
			Public Sub New(ByVal aFullName As String, ByVal anAddress As String)
				Me.mFullName = aFullName
				Me.mAddress = anAddress
				Me.mOrders = New OrderList()
			End Sub

			Public Property FullName() As String
				Get
					Return Me.mFullName
				End Get
				Set(ByVal value As String)
					Me.mFullName = value
				End Set
			End Property

			Public Property Address() As String
				Get
					Return Me.mAddress
				End Get
				Set(ByVal value As String)
					Me.mAddress = value
				End Set
			End Property

			Public Property Orders() As OrderList
				Get
					Return Me.mOrders
				End Get
				Set(ByVal value As OrderList)
					Me.mOrders = value
				End Set
			End Property

			Private mFullName As String
			Private mAddress As String
			Private mOrders As OrderList
		End Class

		''' <summary>
		''' An example of a typed collection that contains your "data" objects.
		''' </summary>
		Public Class CustomerList
			Inherits ArrayList
			Default Public Shadows Property Item(ByVal index As Integer) As Customer
				Get
					Return CType(MyBase.Item(index), Customer)
				End Get
				Set(ByVal value As Customer)
					MyBase.Item(index) = value
				End Set
			End Property
		End Class

		''' <summary>
		''' An example of a child "data entity" class in your application.
		''' </summary>
		Public Class Order
			Public Sub New(ByVal oName As String, ByVal oQuantity As Integer)
				Me.mName = oName
				Me.mQuantity = oQuantity
			End Sub

			Public Property Name() As String
				Get
					Return Me.mName
				End Get
				Set(ByVal value As String)
					Me.mName = value
				End Set
			End Property

			Public Property Quantity() As Integer
				Get
					Return Me.mQuantity
				End Get
				Set(ByVal value As Integer)
					Me.mQuantity = value
				End Set
			End Property

			Private mName As String
			Private mQuantity As Integer
		End Class

		''' <summary>
		''' An example of a typed collection that contains your "data" objects.
		''' </summary>
		Public Class OrderList
			Inherits ArrayList
			Default Public Shadows Property Item(ByVal index As Integer) As Order
				Get
					Return CType(MyBase.Item(index), Order)
				End Get
				Set(ByVal value As Order)
					MyBase.Item(index) = value
				End Set
			End Property
		End Class

		''' <summary>
		''' A custom mail merge data source that you implement to allow Aspose.Words 
		''' to mail merge data from your Customer objects into Microsoft Word documents.
		''' </summary>
		Public Class CustomerMailMergeDataSource
			Implements IMailMergeDataSource
			Public Sub New(ByVal customers As CustomerList)
				Me.mCustomers = customers

				' When the data source is initialized, it must be positioned before the first record.
				Me.mRecordIndex= -1
			End Sub

			''' <summary>
			''' The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
			''' </summary>
            Public ReadOnly Property TableName() As String Implements IMailMergeDataSource.TableName
                Get
                    Return "Customer"
                End Get
            End Property

			''' <summary>
			''' Aspose.Words calls this method to get a value for every data field.
			''' </summary>
            Public Function GetValue(ByVal fieldName As String, <System.Runtime.InteropServices.Out()> ByRef fieldValue As Object) As Boolean Implements IMailMergeDataSource.GetValue
                Select Case fieldName
                    Case "FullName"
                        fieldValue = Me.mCustomers(Me.mRecordIndex).FullName
                        Return True
                    Case "Address"
                        fieldValue = Me.mCustomers(Me.mRecordIndex).Address
                        Return True
                    Case "Order"
                        fieldValue = Me.mCustomers(Me.mRecordIndex).Orders
                        Return True
                    Case Else
                        ' A field with this name was not found, 
                        ' return false to the Aspose.Words mail merge engine.
                        fieldValue = Nothing
                        Return False
                End Select
            End Function


			''' <summary>
			''' A standard implementation for moving to a next record in a collection.
			''' </summary>
            Public Function MoveNext() As Boolean Implements IMailMergeDataSource.MoveNext
                If (Not Me.IsEof) Then
                    Me.mRecordIndex += 1
                End If

                Return ((Not Me.IsEof))
            End Function

			'ExStart
			'ExId:GetChildDataSourceExample
			'ExSummary:Shows how to get a child collection of objects by using the GetChildDataSource method in the parent class.
            Public Function GetChildDataSource(ByVal tableName As String) As IMailMergeDataSource Implements IMailMergeDataSource.GetChildDataSource
                Select Case tableName
                    ' Get the child collection to merge it with the region provided with tableName variable.
                    Case "Order"
                        Return New OrderMailMergeDataSource(Me.mCustomers(Me.mRecordIndex).Orders)
                    Case Else
                        Return Nothing
                End Select
            End Function
			'ExEnd

			Private ReadOnly Property IsEof() As Boolean
				Get
					Return (Me.mRecordIndex >= Me.mCustomers.Count)
				End Get
			End Property

			Private ReadOnly mCustomers As CustomerList
			Private mRecordIndex As Integer
		End Class

		Public Class OrderMailMergeDataSource
			Implements IMailMergeDataSource
			Public Sub New(ByVal orders As OrderList)
				Me.mOrders = orders

				' When the data source is initialized, it must be positioned before the first record.
				Me.mRecordIndex = -1
			End Sub

			''' <summary>
			''' The name of the data source. Used by Aspose.Words only when executing mail merge with repeatable regions.
			''' </summary>
            Public ReadOnly Property TableName() As String Implements IMailMergeDataSource.TableName
                Get
                    Return "Order"
                End Get
            End Property

			''' <summary>
			''' Aspose.Words calls this method to get a value for every data field.
			''' </summary>
            Public Function GetValue(ByVal fieldName As String, <System.Runtime.InteropServices.Out()> ByRef fieldValue As Object) As Boolean Implements IMailMergeDataSource.GetValue
                Select Case fieldName
                    Case "Name"
                        fieldValue = Me.mOrders(Me.mRecordIndex).Name
                        Return True
                    Case "Quantity"
                        fieldValue = Me.mOrders(Me.mRecordIndex).Quantity
                        Return True
                    Case Else
                        ' A field with this name was not found, 
                        ' return false to the Aspose.Words mail merge engine.
                        fieldValue = Nothing
                        Return False
                End Select
            End Function

			''' <summary>
			''' A standard implementation for moving to a next record in a collection.
			''' </summary>
            Public Function MoveNext() As Boolean Implements IMailMergeDataSource.MoveNext
                If (Not Me.IsEof) Then
                    Me.mRecordIndex += 1
                End If

                Return ((Not Me.IsEof))
            End Function

			' Return null because we haven't any child elements for this sort of object.
            Public Function GetChildDataSource(ByVal tableName As String) As IMailMergeDataSource Implements IMailMergeDataSource.GetChildDataSource
                Return Nothing
            End Function

			Private ReadOnly Property IsEof() As Boolean
				Get
					Return (Me.mRecordIndex >= Me.mOrders.Count)
				End Get
			End Property

			Private ReadOnly mOrders As OrderList
			Private mRecordIndex As Integer
		End Class
	End Class
End Namespace
