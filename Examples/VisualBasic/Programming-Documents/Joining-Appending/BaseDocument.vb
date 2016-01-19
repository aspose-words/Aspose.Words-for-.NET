Imports Microsoft.VisualBasic
Imports System
Imports System.Reflection
Imports System.Collections
Imports System.IO
Imports System.Text

Imports Aspose.Words.Lists
Imports Aspose.Words.Fields
Imports Aspose.Words

Public Class BaseDocument
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()
        Dim fileName As String = "TestFile.Destination.doc"

        Dim dstDoc As New Document(dataDir & fileName)
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        ' The destination document is not actually empty which often causes a blank page to appear before the appended document
        ' This is due to the base document having an empty section and the new document being started on the next page.
        ' Remove all content from the destination document before appending.
        dstDoc.RemoveAllChildren()

        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)
        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        dstDoc.Save(dataDir)

        Console.WriteLine(vbNewLine & "Document appended successfully with base document." & vbNewLine & "File saved at " + dataDir)
    End Sub
End Class
