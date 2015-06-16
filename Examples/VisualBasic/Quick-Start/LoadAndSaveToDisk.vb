'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Words

Public Class LoadAndSaveToDisk
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_QuickStart()

        ' Load the document from the absolute path on disk.
        Dim doc As New Document(dataDir & "Document.doc")

        ' Save the document as DOCX document.");
        doc.Save(dataDir & "Document Out.docx")

        Console.WriteLine(vbNewLine + "Existing document loaded and saved successfully." + vbNewLine + "File saved at " + dataDir + "HelloWorld Out.docx")
    End Sub
End Class
