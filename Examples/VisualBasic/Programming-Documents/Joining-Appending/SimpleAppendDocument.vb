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
