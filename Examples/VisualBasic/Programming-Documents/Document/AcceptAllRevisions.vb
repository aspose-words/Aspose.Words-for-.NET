Imports System.IO
Imports Aspose.Words

Public Class AcceptAllRevisions
    Public Shared Sub Run()
        ' ExStart:AcceptAllRevisions
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        Dim doc As New Document(dataDir & Convert.ToString("Document.doc"))

        ' Start tracking and make some revisions.
        doc.StartTrackRevisions("Author")
        doc.FirstSection.Body.AppendParagraph("Hello world!")

        ' Revisions will now show up as normal text in the output document.
        doc.AcceptAllRevisions()

        dataDir = dataDir & Convert.ToString("Document.AcceptedRevisions_out.doc")
        doc.Save(dataDir)
        ' ExEnd:AcceptAllRevisions
        Console.WriteLine(Convert.ToString(vbLf & "All revisions accepted." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
