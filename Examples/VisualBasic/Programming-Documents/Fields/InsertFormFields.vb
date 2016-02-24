Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Fields
Public Class InsertFormFields
    Public Shared Sub Run()
        ' ExStart:InsertFormFields
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()

        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        Dim items As String() = {"One", "Two", "Three"}
        builder.InsertComboBox("DropDown", items, 0)
        ' ExEnd:InsertFormFields
        Console.WriteLine(vbLf & "Form fields inserted successfully.")
    End Sub
End Class
