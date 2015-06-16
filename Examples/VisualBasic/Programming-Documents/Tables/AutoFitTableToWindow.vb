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
Imports System.Diagnostics

Imports Aspose.Words
Imports Aspose.Words.Tables

Public Class AutoFitTableToWindow
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithTables()

        ' Open the document
        Dim doc As New Document(dataDir & "TestFile.doc")

        Dim table As Table = CType(doc.GetChild(NodeType.Table, 0, True), Table)

        ' Autofit the first table to the page width.
        table.AutoFit(AutoFitBehavior.AutoFitToWindow)

        ' Save the document to disk.
        doc.Save(dataDir & "TestFile.AutoFitToWindow Out.doc")
        'ExEnd

        Debug.Assert(doc.FirstSection.Body.Tables(0).PreferredWidth.Type = PreferredWidthType.Percent, "PreferredWidth type is not percent")
        Debug.Assert(doc.FirstSection.Body.Tables(0).PreferredWidth.Value = 100, "PreferredWidth value is different than 100")

        Console.WriteLine(vbNewLine & "Auto fit tables to window successfully." + vbNewLine + "File saved at " + dataDir + "TestFile.AutoFitToWindow Out.doc")
    End Sub
End Class
