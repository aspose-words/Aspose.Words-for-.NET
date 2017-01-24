Imports Microsoft.VisualBasic
Imports System
Imports System.Reflection
Imports System.Collections
Imports System.IO
Imports System.Text

Imports Aspose.Words.Lists
Imports Aspose.Words.Fields
Imports Aspose.Words

Public Class PrependDocument
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()
        Dim fileName As String = "TestFile.Destination.doc"

        Dim dstDoc As New Document(dataDir & fileName)
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        ' Append the source document to the destination document. This causes the result to have line spacing problems.
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)

        ' Instead prepend the content of the destination document to the start of the source document.
        ' This results in the same joined document but with no line spacing issues.
        DoPrepend(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting)

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        ' Save the document
        dstDoc.Save(dataDir)

        Console.WriteLine(vbNewLine & "Document prepended successfully." & vbNewLine & "File saved at " + dataDir)
    End Sub

    Public Shared Sub DoPrepend(dstDoc As Document, srcDoc As Document, mode As ImportFormatMode)
        ' Loop through all sections in the source document. 
        ' Section nodes are immediate children of the Document node so we can just enumerate the Document.
        Dim sections As New ArrayList(srcDoc.Sections.ToArray())

        ' Reverse the order of the sections so they are prepended to start of the destination document in the correct order.
        sections.Reverse()

        For Each srcSection As Section In sections
            ' Import the nodes from the source document.
            Dim dstSection As Node = dstDoc.ImportNode(srcSection, True, mode)

            ' Now the new section node can be prepended to the destination document.
            ' Note how PrependChild is used instead of AppendChild. This is the only line changed compared 
            ' To the original method.
            dstDoc.PrependChild(dstSection)
        Next
    End Sub
End Class
