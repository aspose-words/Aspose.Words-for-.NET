Imports Microsoft.VisualBasic
Imports Aspose.Words
Public Class GetVariables
    Public Shared Sub Run()
        ' ExStart:GetVariables
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        ' Load the template document.
        Dim doc As New Document(dataDir & Convert.ToString("TestFile.doc"))
        Dim variables As String = ""
        For Each entry As DictionaryEntry In doc.Variables
            Dim name As String = entry.Key.ToString()
            Dim value As String = entry.Value.ToString()
            If variables = "" Then
                ' Do something useful.
                variables = Convert.ToString((Convert.ToString("Name: ") & name) + "," + "Value: {1}") & value
            Else
                variables = Convert.ToString((Convert.ToString(variables & Convert.ToString("Name: ")) & name) + "," + "Value: {1}") & value
            End If
        Next
        ' ExEnd:GetVariables
        Console.WriteLine(Convert.ToString(vbLf & "Document have following variables ") & variables)
    End Sub
End Class
