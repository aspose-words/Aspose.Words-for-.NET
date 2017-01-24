Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.IO
Imports System.Reflection
Imports Aspose.Words
Public Class AddComments
    Public Shared Sub Run()
        ' ExStart:AddComments
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithComments()
        ' ExStart:CreateSimpleDocumentUsingDocumentBuilder
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)
        builder.Write("Some text is added.")
        ' ExEnd:CreateSimpleDocumentUsingDocumentBuilder
        Dim comment As New Comment(doc, "Awais Hafeez", "AH", DateTime.Today)
        builder.CurrentParagraph.AppendChild(comment)
        comment.Paragraphs.Add(New Paragraph(doc))
        comment.FirstParagraph.Runs.Add(New Run(doc, "Comment text."))

        dataDir = dataDir & Convert.ToString("Comments_out.doc")
        ' Save the document.
        doc.Save(dataDir)
        ' ExEnd:AddComments
        Console.WriteLine(Convert.ToString(vbLf & "Comments added successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
