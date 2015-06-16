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
Imports Aspose.Words.Fields

Public Class RemoveField
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()

        Dim doc As New Document(dataDir & "Field.RemoveField.doc")

        Dim field As Field = doc.Range.Fields(0)
        ' Calling this method completely removes the field from the document.
        field.Remove()

        Console.WriteLine(vbNewLine & "Field removed from the document successfully." & vbNewLine & "File saved at " + dataDir + "Field.RemoveField.doc")
    End Sub
End Class
