' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_WorkingWithBookmarks()

Dim doc As New Document(dataDir & Convert.ToString("Bookmarks.doc"))

' By index.
Dim bookmark1 As Bookmark = doc.Range.Bookmarks(0)

' By name.
Dim bookmark2 As Bookmark = doc.Range.Bookmarks("Bookmark2")
