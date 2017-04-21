Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports Aspose.Words
Public Class AccessBookmarks
    Public Shared Sub Run()
        ' ExStart:AccessBookmarks
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithBookmarks()

        Dim doc As New Document(dataDir & Convert.ToString("Bookmarks.doc"))

        ' By index.
        Dim bookmark1 As Bookmark = doc.Range.Bookmarks(0)

        ' By name.
        Dim bookmark2 As Bookmark = doc.Range.Bookmarks("Bookmark2")
        ' ExEnd:AccessBookmarks
        Console.WriteLine(vbLf & "Bookmark by name is " + bookmark1.Name + " and bookmark by index is " + bookmark2.Name)
    End Sub
End Class
