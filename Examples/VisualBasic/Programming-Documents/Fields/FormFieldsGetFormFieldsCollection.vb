Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Fields
Public Class FormFieldsGetFormFieldsCollection
    Public Shared Sub Run()
        ' ExStart:FormFieldsGetFormFieldsCollection
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()

        Dim doc As New Document(dataDir & Convert.ToString("FormFields.doc"))
        Dim formFields As FormFieldCollection = doc.Range.FormFields

        ' ExEnd:FormFieldsGetFormFieldsCollection
        Console.WriteLine(vbLf & "Document have " + formFields.Count.ToString() + " form fields.")
    End Sub
End Class
