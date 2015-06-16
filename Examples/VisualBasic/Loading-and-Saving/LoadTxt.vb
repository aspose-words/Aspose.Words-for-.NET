
'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports System.Collections
Imports System.IO

Imports Aspose.Words
Imports Aspose.Words.Tables
Imports System.Diagnostics
Imports Aspose.Words.MailMerging
Imports Aspose.Words.Saving
Imports System.Text

Public Class LoadTxt
    Public Shared Sub Run()

        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()

        ' The encoding of the text file is automatically detected.
        Dim doc As New Document(dataDir & Convert.ToString("LoadTxt.txt"))

        ' Save as any Aspose.Words supported format, such as DOCX.
        doc.Save(dataDir & Convert.ToString("LoadTxt Out.docx"))

        Console.WriteLine(vbNewLine + "Text document loaded successfully." + vbNewLine + "File saved at " + dataDir + "LoadTxt Out.docx")
    End Sub
End Class
