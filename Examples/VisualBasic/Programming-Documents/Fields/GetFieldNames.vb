Imports System.Collections.Generic
Imports Aspose.Words
Imports Aspose.Words.MailMerging

Class GetFieldNames
    Public Shared Sub Run()
        ' ExStart:GetFieldNames
        Dim doc As New Document()
        ' Shows how to get names of all merge fields in a document.
        Dim fieldNames As String() = doc.MailMerge.GetFieldNames()
        ' ExEnd:GetFieldNames
        Console.WriteLine(vbLf & "Document have " & fieldNames.Length.ToString() & " fields.")
    End Sub
    Public Sub MappedDataFields()
        ' ExStart:MappedDataFields
        Dim doc As New Document()
        ' Shows how to add a mapping when a merge field in a document and a data field in a data source have different names.
        doc.MailMerge.MappedDataFields.Add("MyFieldName_InDocument", "MyFieldName_InDataSource")
        ' ExEnd:MappedDataFields
    End Sub
    Public Sub DeleteFields()
        ' ExStart:DeleteFields
        Dim doc As New Document()
        ' Shows how to delete all merge fields from a document without executing mail merge.
        doc.MailMerge.DeleteFields()
        ' ExEnd:DeleteFields
    End Sub
End Class
