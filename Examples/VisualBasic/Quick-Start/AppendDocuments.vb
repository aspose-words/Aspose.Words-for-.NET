Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Words

Public Class AppendDocuments
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_QuickStart()

        ' Load the destination and source documents from disk.
        Dim dstDoc As New Document(dataDir & "TestFile.Destination.doc")
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        ' Append the source document to the destination document while keeping the original formatting of the source document.
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)

        dstDoc.Save(dataDir & "TestFile Out.docx")

        Console.WriteLine(vbNewLine + "Document appended successfully." + vbNewLine + "File saved at " + dataDir + "TestFile Out.docx")
    End Sub
End Class
