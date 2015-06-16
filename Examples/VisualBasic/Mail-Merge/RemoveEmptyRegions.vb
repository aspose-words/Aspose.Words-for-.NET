﻿'////////////////////////////////////////////////////////////////////////
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
Imports System.Data
Imports System.Diagnostics

Imports Aspose.Words
Imports Aspose.Words.Reporting
Imports Aspose.Words.MailMerging

Public Class RemoveEmptyRegions
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()

        ' Open the document.
        Dim doc As New Document(dataDir & "TestFile.doc")

        ' Create a dummy data source containing no data.
        Dim data As New DataSet()

        ' Set the appropriate mail merge clean up options to remove any unused regions from the document.
        doc.MailMerge.CleanupOptions = MailMergeCleanupOptions.RemoveUnusedRegions

        ' Execute mail merge which will have no effect as there is no data. However the regions found in the document will be removed
        ' automatically as they are unused.
        doc.MailMerge.ExecuteWithRegions(data)

        ' Save the output document to disk.
        doc.Save(dataDir & "TestFile.RemoveEmptyRegions Out.doc")

        Console.WriteLine(vbNewLine + "Mail merge performed with empty regions successfully." + vbNewLine + "File saved at " + dataDir + "TestFile.RemoveEmptyRegions Out.doc")
    End Sub
End Class
