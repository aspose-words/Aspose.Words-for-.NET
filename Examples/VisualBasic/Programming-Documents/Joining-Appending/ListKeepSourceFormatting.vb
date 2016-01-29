Imports Microsoft.VisualBasic
Imports System
Imports System.Reflection
Imports System.Collections
Imports System.IO
Imports System.Text

Imports Aspose.Words.Lists
Imports Aspose.Words.Fields
Imports Aspose.Words

Public Class ListKeepSourceFormatting
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()
        Dim fileName As String = "TestFile.Destination.doc"

        Dim dstDoc As New Document(dataDir & fileName)
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        ' Append the content of the document so it flows continuously.
        srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous

        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        dstDoc.Save(dataDir)

        Console.WriteLine(vbNewLine & "Document appended successfully with lists keeping source formatting." & vbNewLine & "File saved at " + dataDir)
    End Sub
End Class
