Imports Microsoft.VisualBasic
Imports System
Imports System.Reflection
Imports System.Collections
Imports System.IO
Imports System.Text

Imports Aspose.Words.Lists
Imports Aspose.Words.Fields
Imports Aspose.Words

Public Class JoinNewPage
    Public Shared Sub Run()
        ' ExStart:JoinNewPage
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()
        Dim fileName As String = "TestFile.Destination.doc"

        Dim dstDoc As New Document(dataDir & fileName)
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        ' Set the appended document to start on a new page.
        srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        ' Append the source document using the original styles found in the source document.
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)
        dstDoc.Save(dataDir)
        ' ExEnd:JoinNewPage
        Console.WriteLine(vbNewLine & "Document appended successfully with join new page option." & vbNewLine & "File saved at " + dataDir)
    End Sub
End Class
