Imports Microsoft.VisualBasic
Imports System
Imports System.Reflection
Imports System.Collections
Imports System.IO
Imports System.Text

Imports Aspose.Words.Lists
Imports Aspose.Words.Fields
Imports Aspose.Words

Public Class LinkHeadersFooters
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()
        Dim fileName As String = "TestFile.Destination.doc"

        Dim dstDoc As New Document(dataDir & fileName)
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        ' Set the appended document to appear on a new page.
        srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage

        ' Link the headers and footers in the source document to the previous section. 
        ' This will override any headers or footers already found in the source document. 
        srcDoc.FirstSection.HeadersFooters.LinkToPrevious(True)

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)
        dstDoc.Save(dataDir)

        Console.WriteLine(vbNewLine & "Document appended successfully with linked header footers." & vbNewLine & "File saved at " + dataDir)
    End Sub
End Class
