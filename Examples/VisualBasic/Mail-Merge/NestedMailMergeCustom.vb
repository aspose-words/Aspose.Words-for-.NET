Imports System.Collections
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Tables
Imports System.Diagnostics
Imports Aspose.Words.MailMerging
Imports Aspose.Words.Saving
Imports System.Text

Class NestedMailMergeCustom
    Public Shared Sub Run()
        ' ExStart:NestedMailMergeCustom
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()
        Dim fileName As String = "NestedMailMerge.CustomDataSource.doc"
        ' Create some data that we will use in the mail merge.
        Dim customers As New CustomerList()
        customers.Add(New Customer("Thomas Hardy", "120 Hanover Sq., London"))
        customers.Add(New Customer("Paolo Accorti", "Via Monte Bianco 34, Torino"))

        ' Create some data for nesting in the mail merge.
        customers(0).Orders.Add(New Order("Rugby World Cup Cap", 2))
        customers(0).Orders.Add(New Order("Rugby World Cup Ball", 1))
        customers(1).Orders.Add(New Order("Rugby World Cup Guide", 1))

        ' Open the template document.
        Dim doc As New Document(dataDir & fileName)

        ' To be able to mail merge from your own data source, it must be wrapped
        ' into an object that implements the IMailMergeDataSource interface.
        Dim customersDataSource As New CustomerMailMergeDataSource(customers)

        ' Now you can pass your data source into Aspose.Words.
        doc.MailMerge.ExecuteWithRegions(customersDataSource)

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        doc.Save(dataDir)
        ' ExEnd:NestedMailMergeCustom

        Console.WriteLine(Convert.ToString(vbLf & "Mail merge performed with nested custom data successfully." & vbLf & "File saved at ") & dataDir)
    End Sub

    ''' <summary>
    ''' An example of a "data entity" class in your application.
    ''' </summary>
    Public Class Customer
        Public Sub New(aFullName As String, anAddress As String)
            mFullName = aFullName
            mAddress = anAddress
            mOrders = New OrderList()
        End Sub

        Public Property FullName() As String
            Get
                Return mFullName
            End Get
            Set(value As String)
                mFullName = value
            End Set
        End Property

        Public Property Address() As String
            Get
                Return mAddress
            End Get
            Set(value As String)
                mAddress = value
            End Set
        End Property

        Public Property Orders() As OrderList
            Get
                Return mOrders
            End Get
            Set(value As OrderList)
                mOrders = value
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
        Default Public Shadows Property Item(index As Integer) As Customer
            Get
                Return DirectCast(MyBase.Item(index), Customer)
            End Get
            Set(value As Customer)
                MyBase.Item(index) = value
            End Set
        End Property
    End Class

    ''' <summary>
    ''' An example of a child "data entity" class in your application.
    ''' </summary>
    Public Class Order
        Public Sub New(oName As String, oQuantity As Integer)
            mName = oName
            mQuantity = oQuantity
        End Sub

        Public Property Name() As String
            Get
                Return mName
            End Get
            Set(value As String)
                mName = value
            End Set
        End Property

        Public Property Quantity() As Integer
            Get
                Return mQuantity
            End Get
            Set(value As Integer)
                mQuantity = value
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
        Default Public Shadows Property Item(index As Integer) As Order
            Get
                Return DirectCast(MyBase.Item(index), Order)
            End Get
            Set(value As Order)
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
            mCustomers = customers

            ' When the data source is initialized, it must be positioned before the first record.
            mRecordIndex = -1
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
                    fieldValue = mCustomers(mRecordIndex).FullName
                    Return True
                Case "Address"
                    fieldValue = mCustomers(mRecordIndex).Address
                    Return True
                Case "Order"
                    fieldValue = mCustomers(mRecordIndex).Orders
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
            If (Not IsEof) Then
                mRecordIndex += 1
            End If

            Return ((Not IsEof))
        End Function

        ' ExStart:GetChildDataSourceExample
        Public Function GetChildDataSource(ByVal tableName As String) As IMailMergeDataSource Implements IMailMergeDataSource.GetChildDataSource
            Select Case tableName
                ' Get the child collection to merge it with the region provided with tableName variable.
                Case "Order"
                    Return New OrderMailMergeDataSource(mCustomers(mRecordIndex).Orders)
                Case Else
                    Return Nothing
            End Select
        End Function
        ' ExEnd:GetChildDataSourceExample

        Private ReadOnly Property IsEof() As Boolean
            Get
                Return (mRecordIndex >= mCustomers.Count)
            End Get
        End Property

        Private ReadOnly mCustomers As CustomerList
        Private mRecordIndex As Integer
    End Class

    Public Class OrderMailMergeDataSource
        Implements IMailMergeDataSource
        Public Sub New(ByVal orders As OrderList)
            mOrders = orders

            ' When the data source is initialized, it must be positioned before the first record.
            mRecordIndex = -1
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
                    fieldValue = mOrders(mRecordIndex).Name
                    Return True
                Case "Quantity"
                    fieldValue = mOrders(mRecordIndex).Quantity
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
            If (Not IsEof) Then
                mRecordIndex += 1
            End If

            Return ((Not IsEof))
        End Function

        ' Return null because we haven't any child elements for this sort of object.
        Public Function GetChildDataSource(ByVal tableName As String) As IMailMergeDataSource Implements IMailMergeDataSource.GetChildDataSource
            Return Nothing
        End Function

        Private ReadOnly Property IsEof() As Boolean
            Get
                Return (mRecordIndex >= mOrders.Count)
            End Get
        End Property

        Private ReadOnly mOrders As OrderList
        Private mRecordIndex As Integer
    End Class
End Class
