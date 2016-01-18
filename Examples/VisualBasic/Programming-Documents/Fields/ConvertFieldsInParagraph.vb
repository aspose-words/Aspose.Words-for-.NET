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

        Dim doc As New Document(dataDir & "TestFile.doc")

        ' Pass the appropriate parameters to convert all IF fields to static text that are encountered only in the last 
        ' paragraph of the document.
        FieldsHelper.ConvertFieldsToStaticText(doc.FirstSection.Body.LastParagraph, FieldType.FieldIf)

        ' Save the document with fields transformed to disk.
        doc.Save(dataDir & "TestFileParagraph Out.doc")

        Console.WriteLine(vbNewLine & "Converted fields to static text in the paragraph successfully." & vbNewLine & "File saved at " + dataDir + "TestFileParagraph Out.doc")
    End Sub
End Class
