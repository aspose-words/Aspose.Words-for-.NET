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

Public Class AppendDocumentManually
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()

        Dim dstDoc As New Document(dataDir & "TestFile.Destination.doc")
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")


        Dim mode As ImportFormatMode = ImportFormatMode.KeepSourceFormatting

        ' Loop through all sections in the source document. 
        ' Section nodes are immediate children of the Document node so we can just enumerate the Document.
        For Each srcSection As Section In srcDoc
            ' Because we are copying a section from one document to another, 
            ' it is required to import the Section node into the destination document.
            ' This adjusts any document-specific references to styles, lists, etc.
            '
            ' Importing a node creates a copy of the original node, but the copy
            ' is ready to be inserted into the destination document.
            Dim dstSection As Node = dstDoc.ImportNode(srcSection, True, mode)

            ' Now the new section node can be appended to the destination document.
            dstDoc.AppendChild(dstSection)
        Next

        ' Save the joined document
        dstDoc.Save(dataDir + "TestFile.Append Manual Out.doc")

        Console.WriteLine(vbNewLine & "Document appended successfully with updated page layout." & vbNewLine & "File saved at " + dataDir + "TestFile.Append Manual Out.docx")
    End Sub
End Class
