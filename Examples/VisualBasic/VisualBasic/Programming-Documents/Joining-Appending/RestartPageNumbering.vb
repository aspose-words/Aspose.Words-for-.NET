Imports Microsoft.VisualBasic
Imports System
Imports System.Reflection
Imports System.Collections
Imports System.IO
Imports System.Text

Imports Aspose.Words.Lists
Imports Aspose.Words.Fields
Imports Aspose.Words

Public Class RestartPageNumbering
    Public Shared Sub Run()
        ' ExStart:RestartPageNumbering
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()
        Dim fileName As String = "TestFile.Destination.doc"

        Dim dstDoc As New Document(dataDir & fileName)
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        ' Set the appended document to appear on the next page.
        srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage
        ' Restart the page numbering for the document to be appended.
        srcDoc.FirstSection.PageSetup.RestartPageNumbering = True

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)
        dstDoc.Save(dataDir)
        ' ExEnd:RestartPageNumbering
        Console.WriteLine(vbNewLine & "Document appended successfully with restart page numbering option." & vbNewLine & "File saved at " + dataDir)
    End Sub
End Class
