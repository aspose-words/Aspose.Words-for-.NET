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

Imports Aspose.Words

Public Class FindAndReplace
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_QuickStart()

        ' Open the document.
        Dim doc As New Document(dataDir & "ReplaceSimple.doc")

        ' Check the text of the document
        Console.WriteLine("Original document text: " & doc.Range.Text)

        ' Replace the text in the document.
        doc.Range.Replace("_CustomerName_", "James Bond", False, False)

        ' Check the replacement was made.
        Console.WriteLine("Document text after replace: " & doc.Range.Text)

        ' Save the modified document.
        doc.Save(dataDir & "ReplaceSimple Out.doc")

        Console.WriteLine(vbNewLine + "Text found and replaced successfully." + vbNewLine + "File saved at " + dataDir + "ReplaceSimple Out.doc")
    End Sub
End Class
