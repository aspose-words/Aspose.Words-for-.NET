Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Words

Public Class Doc2Pdf
    Public Shared Sub Run()
        ' ExStart:Doc2Pdf
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_QuickStart()

        ' Load the document from disk.
        Dim doc As New Document(dataDir & "Template.doc")

        dataDir = dataDir & "Template_out_.pdf"
        ' Save the document in PDF format.
        doc.Save(dataDir)
        ' ExEnd:Doc2Pdf

        Console.WriteLine(vbNewLine + "Document converted to PDF successfully." + vbNewLine + "File saved at " + dataDir)
    End Sub
End Class
