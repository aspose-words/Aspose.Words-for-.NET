Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection
Imports System.Collections

Imports Aspose.Words
Imports Aspose.Words.Fields
Imports Aspose.Words.Tables
Imports System.Diagnostics

Public Class ConvertFieldsInParagraph
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()
        Dim fileName As String = "TestFile.doc"
        Dim doc As New Document(dataDir & fileName)

        ' Pass the appropriate parameters to convert all IF fields to static text that are encountered only in the last 
        ' paragraph of the document.
        FieldsHelper.ConvertFieldsToStaticText(doc.FirstSection.Body.LastParagraph, FieldType.FieldIf)

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        ' Save the document with fields transformed to disk.
        doc.Save(dataDir)

        Console.WriteLine(vbNewLine & "Converted fields to static text in the paragraph successfully." & vbNewLine & "File saved at " + dataDir)
    End Sub
End Class
