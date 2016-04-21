Imports System
Imports System.Collections
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Tables
Imports Aspose.Words.Fields
Imports Aspose.Words.Layout
Imports System.Text.RegularExpressions
Imports System.Text
Public Class ReplaceHyperlinks
    Public Shared Sub Run()
        ' ExStart:ReplaceHyperlinks
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithHyperlink()
        Dim NewUrl As String = "http://www.aspose.com"
        Dim NewName As String = "Aspose - The .NET & Java Component Publisher"
        Dim doc As New Document(dataDir & Convert.ToString("ReplaceHyperlinks.doc"))

        ' Hyperlinks in a Word documents are fields.
        For Each field As Field In doc.Range.Fields
            If field.Type = FieldType.FieldHyperlink Then
                Dim hyperlink As FieldHyperlink = DirectCast(field, FieldHyperlink)

                ' Some hyperlinks can be local (links to bookmarks inside the document), ignore these.
                If hyperlink.SubAddress IsNot Nothing Then
                    Continue For
                End If

                hyperlink.Address = NewUrl
                hyperlink.Result = NewName
            End If
        Next
        dataDir = dataDir & Convert.ToString("ReplaceHyperlinks_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:ReplaceHyperlinks
        Console.WriteLine(Convert.ToString(vbLf & "Hyperlinks replaced successfully." & vbLf & "File saved at ") & dataDir)
    End Sub

End Class
