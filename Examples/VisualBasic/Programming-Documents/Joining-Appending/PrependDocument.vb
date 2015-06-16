'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

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

        Dim dstDoc As New Document(dataDir & "TestFile.Destination.doc")
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        ' Append the source document to the destination document. This causes the result to have line spacing problems.
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)

        ' Instead prepend the content of the destination document to the start of the source document.
        ' This results in the same joined document but with no line spacing issues.
        DoPrepend(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting)

        ' Save the document
        dstDoc.Save(dataDir + "TestFile.Prepend.doc")

        Console.WriteLine(vbNewLine & "Document prepended successfully." & vbNewLine & "File saved at " + dataDir + "TestFile.Prepend Out.docx")
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
            ' to the original method.
            dstDoc.PrependChild(dstSection)
        Next
    End Sub
End Class
