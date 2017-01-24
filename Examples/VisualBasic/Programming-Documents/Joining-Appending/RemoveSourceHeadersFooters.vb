Imports Microsoft.VisualBasic
Imports System
Imports System.Reflection
Imports System.Collections
Imports System.IO
Imports System.Text

Imports Aspose.Words.Lists
Imports Aspose.Words.Fields
Imports Aspose.Words

Public Class RemoveSourceHeadersFooters
    Public Shared Sub Run()
        ' ExStart:RemoveSourceHeadersFooters
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()
        Dim fileName As String = "TestFile.Destination.doc"
        Dim dstDoc As New Document(dataDir & fileName)
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        ' Remove the headers and footers from each of the sections in the source document.
        For Each section As Section In srcDoc.Sections
            section.ClearHeadersFooters()
        Next section

        ' Even after the headers and footers are cleared from the source document, the "LinkToPrevious" setting 
        ' For HeadersFooters can still be set. This will cause the headers and footers to continue from the destination 
        ' Document. This should set to false to avoid this behavior.
        srcDoc.FirstSection.HeadersFooters.LinkToPrevious(False)

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)
        dstDoc.Save(dataDir)
        ' ExEnd:RemoveSourceHeadersFooters
        Console.WriteLine(vbNewLine & "Document appended successfully with removed source header footers." & vbNewLine & "File saved at " + dataDir)
    End Sub
End Class
