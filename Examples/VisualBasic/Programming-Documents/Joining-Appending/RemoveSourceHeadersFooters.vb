﻿'////////////////////////////////////////////////////////////////////////
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

Public Class RemoveSourceHeadersFooters
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()

        Dim dstDoc As New Document(dataDir & "TestFile.Destination.doc")
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        ' Remove the headers and footers from each of the sections in the source document.
        For Each section As Section In srcDoc.Sections
            section.ClearHeadersFooters()
        Next section

        ' Even after the headers and footers are cleared from the source document, the "LinkToPrevious" setting 
        ' for HeadersFooters can still be set. This will cause the headers and footers to continue from the destination 
        ' document. This should set to false to avoid this behavior.
        srcDoc.FirstSection.HeadersFooters.LinkToPrevious(False)

        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)
        dstDoc.Save(dataDir & "TestFile.RemoveSourceHeadersFooters Out.doc")

        Console.WriteLine(vbNewLine & "Document appended successfully with removed source header footers." & vbNewLine & "File saved at " + dataDir + "TestFile.RemoveSourceHeadersFooters Out.docx")
    End Sub
End Class
