Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Saving
Public Class CreateBookmark
    Public Shared Sub Run()
        ' ExStart:CreateBookmark
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithBookmarks()

        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        builder.StartBookmark("My Bookmark")
        builder.Writeln("Text inside a bookmark.")

        builder.StartBookmark("Nested Bookmark")
        builder.Writeln("Text inside a NestedBookmark.")
        builder.EndBookmark("Nested Bookmark")

        builder.Writeln("Text after Nested Bookmark.")
        builder.EndBookmark("My Bookmark")


        Dim options As New PdfSaveOptions()
        options.OutlineOptions.BookmarksOutlineLevels.Add("My Bookmark", 1)
        options.OutlineOptions.BookmarksOutlineLevels.Add("Nested Bookmark", 2)

        dataDir = dataDir & Convert.ToString("Create.Bookmark_out.pdf")
        doc.Save(dataDir, options)
        ' ExEnd:CreateBookmark
        Console.WriteLine(Convert.ToString(vbLf & "Bookmark created successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
