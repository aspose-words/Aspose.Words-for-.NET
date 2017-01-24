' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports System
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic
Imports System.Collections

Imports Aspose.Words
Imports Aspose.Words.MailMerging

Imports NUnit.Framework

Namespace ApiExamples
    <TestFixture> _
    Public Class ExMailMergeCustom
        Inherits ApiExampleBase
        ''' <summary>
        ''' This calls the below method to resolve skipping of [Test] in VB.NET.
        ''' </summary>
        <Test> _
        Public Sub MailMergeCustomDataSourceCaller()
            Me.MailMergeCustomDataSource()
        End Sub

        'ExStart
        'ExFor:IMailMergeDataSource
        'ExFor:IMailMergeDataSource.TableName
        'ExFor:IMailMergeDataSource.MoveNext
        'ExFor:IMailMergeDataSource.GetValue
        'ExFor:IMailMergeDataSource.GetChildDataSource
        'ExFor:MailMerge.Execute(IMailMergeDataSource)
        'ExSummary:Performs mail merge from a custom data source.
        Public Sub MailMergeCustomDataSource()
            ' Create some data that we will use in the mail merge.
            Dim customers As New CustomerList()
            customers.Add(New Customer("Thomas Hardy", "120 Hanover Sq., London"))
            customers.Add(New Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"))

            ' Open the template document.
            Dim doc As New Document(MyDir & "MailMerge.CustomDataSource.doc")

            ' To be able to mail merge from your own data source, it must be wrapped
            ' into an object that implements the IMailMergeDataSource interface.
            Dim customersDataSource As New CustomerMailMergeDataSource(customers)

            ' Now you can pass your data source into Aspose.Words.
            doc.MailMerge.Execute(customersDataSource)

            doc.Save(MyDir & "\Artifacts\MailMerge.CustomDataSource.doc")
        End Sub

        ''' <summary>
        ''' An example of a "data entity" class in your application.
        ''' </summary>
        Public Class Customer
            Public Sub New(ByVal aFullName As String, ByVal anAddress As String)
                Me.mFullName = aFullName
                Me.mAddress = anAddress
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

            Private mFullName As String
            Private mAddress As String
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
        ''' A custom mail merge data source that you implement to allow Aspose.Words 
        ''' to mail merge data from your Customer objects into Microsoft Word documents.
        ''' </summary>
        Public Class CustomerMailMergeDataSource
            Implements IMailMergeDataSource
            Public Sub New(ByVal customers As CustomerList)
                Me.mCustomers = customers

                ' When the data source is initialized, it must be positioned before the first record.
                Me.mRecordIndex = -1
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

            Public Function GetChildDataSource(ByVal tableName As String) As IMailMergeDataSource Implements IMailMergeDataSource.GetChildDataSource
                Return Nothing
            End Function

            Private ReadOnly Property IsEof() As Boolean
                Get
                    Return (Me.mRecordIndex >= Me.mCustomers.Count)
                End Get
            End Property

            Private ReadOnly mCustomers As CustomerList
            Private mRecordIndex As Integer
        End Class
        'ExEnd
    End Class
End Namespace
