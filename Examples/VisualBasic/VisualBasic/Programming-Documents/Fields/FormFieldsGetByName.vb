Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Fields
Public Class FormFieldsGetByName
    Public Shared Sub Run()
        ' ExStart:FormFieldsGetByName
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()

        Dim doc As New Document(dataDir & Convert.ToString("FormFields.doc"))
        Dim documentFormFields As FormFieldCollection = doc.Range.FormFields

        Dim formField1 As FormField = documentFormFields(3)
        Dim formField2 As FormField = documentFormFields("Text2")
        ' ExEnd:FormFieldsGetByName
        Console.WriteLine(vbLf + formField2.Name + " field have following text " + formField2.GetText() + ".")
    End Sub
End Class
