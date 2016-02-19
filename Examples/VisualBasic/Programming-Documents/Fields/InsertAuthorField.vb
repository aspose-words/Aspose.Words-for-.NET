Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Fields
Public Class InsertAuthorField
    Public Shared Sub Run()
        ' ExStart:InsertAuthorField
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()
        Dim doc As New Document(dataDir & Convert.ToString("in.doc"))
        ' Get paragraph you want to append this AUTHOR field to
        Dim para As Paragraph = DirectCast(doc.GetChildNodes(NodeType.Paragraph, True)(1), Paragraph)

        ' We want to insert an AUTHOR field like this:
        ' { AUTHOR Test1 }

        ' Create instance of FieldAuthor class and lets build the above field code
        Dim field As FieldAuthor = DirectCast(para.AppendField(FieldType.FieldAuthor, False), FieldAuthor)

        ' { AUTHOR Test1 }
        field.AuthorName = "Test1"

        ' Finally update this AUTHOR field
        field.Update()

        dataDir = dataDir & Convert.ToString("InsertAuthorField_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:InsertAuthorField
        Console.WriteLine(Convert.ToString(vbLf & "Author field without document builder inserted successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
