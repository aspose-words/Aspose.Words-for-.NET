Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Words

Public Class Doc2Pdf
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_QuickStart()

        ' Load the document from disk.
        Dim doc As New Document(dataDir & "Template.doc")

        ' Save the document in PDF format.
        doc.Save(dataDir & "Doc2PdfSave Out.pdf")

        Console.WriteLine(vbNewLine + "Document converted to PDF successfully." + vbNewLine + "File saved at " + dataDir + "Doc2PdfSave Out.pdf")
    End Sub
End Class
