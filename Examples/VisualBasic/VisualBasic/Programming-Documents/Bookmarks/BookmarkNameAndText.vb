Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports Aspose.Words
Public Class BookmarkNameAndText
    Public Shared Sub Run()
        ' ExStart:BookmarkNameAndText
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithBookmarks()

        Dim doc As New Document(dataDir & Convert.ToString("Bookmark.doc"))

        ' Use the indexer of the Bookmarks collection to obtain the desired bookmark.
        Dim bookmark As Bookmark = doc.Range.Bookmarks("MyBookmark")

        ' Get the name and text of the bookmark.
        Dim name As String = bookmark.Name
        Dim text As String = bookmark.Text

        ' Set the name and text of the bookmark.
        bookmark.Name = "RenamedBookmark"
        bookmark.Text = "This is a new bookmarked text."
        ' ExEnd:BookmarkNameAndText
        Console.WriteLine(Convert.ToString(vbLf & "Bookmark text and name get and set successfully.") & dataDir)
    End Sub
End Class
