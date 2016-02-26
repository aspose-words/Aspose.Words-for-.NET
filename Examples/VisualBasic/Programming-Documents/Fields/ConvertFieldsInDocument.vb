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
        ' ExStart:ConvertFieldsInDocument
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()
        Dim fileName As String = "TestFile.doc"

        Dim doc As New Document(dataDir & fileName)

        ' Pass the appropriate parameters to convert all IF fields encountered in the document (including headers and footers) to static text.
        FieldsHelper.ConvertFieldsToStaticText(doc, FieldType.FieldIf)

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        ' Save the document with fields transformed to disk.
        doc.Save(dataDir)
        ' ExEnd:ConvertFieldsInDocument
        Console.WriteLine(vbNewLine & "Converted fields to static text in the document successfully." & vbNewLine & "File saved at " + dataDir)
    End Sub
End Class
