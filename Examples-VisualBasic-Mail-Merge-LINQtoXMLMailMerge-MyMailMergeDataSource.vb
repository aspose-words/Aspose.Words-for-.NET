' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Public Class MyMailMergeDataSource
    Implements IMailMergeDataSource
    
    Public Sub New(ByVal data As IEnumerable)
        mEnumerator = data.GetEnumerator()
    End Sub
    
    Public Sub New(ByVal data As IEnumerable, ByVal tableName As String)
        mEnumerator = data.GetEnumerator()
        mTableName = tableName
    End Sub

    Public Function GetValue(ByVal fieldName As String, <System.Runtime.InteropServices.Out()> ByRef fieldValue As Object) As Boolean Implements IMailMergeDataSource.GetValue
        ' Use reflection to get the property by name from the current object.
        Dim obj As Object = mEnumerator.Current

        Dim curentRecordType As Type = obj.GetType()
        Dim [property] As PropertyInfo = curentRecordType.GetProperty(fieldName)
        If [property] IsNot Nothing Then
            fieldValue = [property].GetValue(obj, Nothing)
            Return True
        End If

        ' Return False to the Aspose.Words mail merge engine to indicate the field was not found.
        fieldValue = Nothing
        Return False
    End Function
    
    Public Function MoveNext() As Boolean Implements IMailMergeDataSource.MoveNext
        Return mEnumerator.MoveNext()
    End Function
    
    Public ReadOnly Property TableName() As String Implements IMailMergeDataSource.TableName
        Get
            Return mTableName
        End Get
    End Property
    
    Public Function GetChildDataSource(ByVal tableName As String) As IMailMergeDataSource Implements IMailMergeDataSource.GetChildDataSource
        Return Nothing
    End Function

    Private ReadOnly mEnumerator As IEnumerator
    Private ReadOnly mTableName As String
End Class
