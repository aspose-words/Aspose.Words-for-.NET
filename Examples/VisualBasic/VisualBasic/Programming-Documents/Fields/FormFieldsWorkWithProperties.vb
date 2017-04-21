Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Fields
Public Class FormFieldsWorkWithProperties
    Public Shared Sub Run()
        ' ExStart:FormFieldsWorkWithProperties
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()

        Dim doc As New Document(dataDir & Convert.ToString("FormFields.doc"))
        Dim formField As FormField = doc.Range.FormFields(3)

        If formField.Type.Equals(FieldType.FieldFormTextInput) Then
            formField.Result = "My name is " + formField.Name
        End If
        ' ExEnd:FormFieldsWorkWithProperties            
    End Sub
End Class
