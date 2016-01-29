Imports Microsoft.VisualBasic
Imports System
Imports System.Reflection
Imports System.Collections
Imports System.IO
Imports System.Text

Imports Aspose.Words.Lists
Imports Aspose.Words.Fields
Imports Aspose.Words

Public Class UnlinkHeadersFooters
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()
        Dim fileName As String = "TestFile.Destination.doc"

        Dim dstDoc As New Document(dataDir & fileName)
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        ' Even a document with no headers or footers can still have the LinkToPrevious setting set to true.
        ' Unlink the headers and footers in the source document to stop this from continuing the headers and footers
        ' from the destination document.
        srcDoc.FirstSection.HeadersFooters.LinkToPrevious(False)

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)
        dstDoc.Save(dataDir)

        Console.WriteLine(vbNewLine & "Document appended successfully with un-linked header footers." & vbNewLine & "File saved at " + dataDir)
    End Sub
End Class
