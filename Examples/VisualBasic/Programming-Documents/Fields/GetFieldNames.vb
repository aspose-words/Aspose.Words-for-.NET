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
End Class
