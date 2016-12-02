Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Fields
Public Class InsertField
    Public Shared Sub Run()
        ' ExStart:InsertField
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()

        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)
        builder.InsertField("MERGEFIELD MyFieldName \* MERGEFORMAT")
        dataDir = dataDir & Convert.ToString("InsertField_out.docx")
        doc.Save(dataDir)
        ' ExEnd:InsertField
        Console.WriteLine(Convert.ToString(vbLf & "Inserted field in the document successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
