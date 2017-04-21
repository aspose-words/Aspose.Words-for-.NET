Imports Microsoft.VisualBasic
Imports Aspose.Words
Public Class AccessStyles
    Public Shared Sub Run()
        ' ExStart:AccessStyles
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        ' Load the template document.
        Dim doc As New Document(dataDir & Convert.ToString("TestFile.doc"))
        ' Get styles collection from document.
        Dim styles As StyleCollection = doc.Styles
        Dim styleName As String = ""
        ' Iterate through all the styles.
        For Each style As Style In styles
            If styleName = "" Then
                styleName = style.Name
            Else
                styleName = (styleName & Convert.ToString(", ")) + style.Name
            End If
        Next
        ' ExEnd:AccessStyles
        Console.WriteLine(Convert.ToString(vbLf & "Document have following styles ") & styleName)
    End Sub
End Class
