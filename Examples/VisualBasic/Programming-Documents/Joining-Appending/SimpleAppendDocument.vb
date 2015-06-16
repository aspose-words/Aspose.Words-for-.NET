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

Public Class SimpleAppendDocument
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()

        Dim dstDoc As New Document(dataDir & "TestFile.Destination.doc")
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)
        
        dstDoc.Save(dataDir & "TestFile.SimpleAppendDocument Out.docx")

        Console.WriteLine(vbNewLine & "Simple document append." & vbNewLine & "File saved at " + dataDir + "TestFile.SimpleAppendDocument Out.docx")
    End Sub
End Class
