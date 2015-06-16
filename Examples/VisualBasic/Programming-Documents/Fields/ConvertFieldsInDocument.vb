'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection
Imports System.Collections

Imports Aspose.Words
Imports Aspose.Words.Fields
Imports Aspose.Words.Tables
Imports System.Diagnostics

Public Class ConvertFieldsInDocument
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()

        Dim doc As New Document(dataDir & "TestFile.doc")

        ' Pass the appropriate parameters to convert all IF fields encountered in the document (including headers and footers) to static text.
        FieldsHelper.ConvertFieldsToStaticText(doc, FieldType.FieldIf)

        ' Save the document with fields transformed to disk.
        doc.Save(dataDir & "TestFileDocument Out.doc")

        Console.WriteLine(vbNewLine & "Converted fields to static text in the document successfully." & vbNewLine & "File saved at " + dataDir + "TestFileDocument Out.doc")
    End Sub
End Class
