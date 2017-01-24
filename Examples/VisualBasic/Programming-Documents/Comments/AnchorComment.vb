Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.IO
Imports System.Reflection
Imports Aspose.Words
Public Class AnchorComment
    Public Shared Sub Run()
        ' ExStart:AnchorComment
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithComments()
        Dim doc As New Document()

        Dim para1 As New Paragraph(doc)
        Dim run1 As New Run(doc, "Some ")
        Dim run2 As New Run(doc, "text ")
        para1.AppendChild(run1)
        para1.AppendChild(run2)
        doc.FirstSection.Body.AppendChild(para1)

        Dim para2 As New Paragraph(doc)
        Dim run3 As New Run(doc, "is ")
        Dim run4 As New Run(doc, "added ")
        para2.AppendChild(run3)
        para2.AppendChild(run4)
        doc.FirstSection.Body.AppendChild(para2)

        Dim comment As New Comment(doc, "Awais Hafeez", "AH", DateTime.Today)
        comment.Paragraphs.Add(New Paragraph(doc))
        comment.FirstParagraph.Runs.Add(New Run(doc, "Comment text."))

        Dim commentRangeStart As New CommentRangeStart(doc, comment.Id)
        Dim commentRangeEnd As New CommentRangeEnd(doc, comment.Id)

        run1.ParentNode.InsertAfter(commentRangeStart, run1)
        run3.ParentNode.InsertAfter(commentRangeEnd, run3)
        commentRangeEnd.ParentNode.InsertAfter(comment, commentRangeEnd)

        dataDir = dataDir & Convert.ToString("Anchor.Comment_out.doc")
        ' Save the document.
        doc.Save(dataDir)
        ' ExEnd:AnchorComment
        Console.WriteLine(Convert.ToString(vbLf & "Comment anchored successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
