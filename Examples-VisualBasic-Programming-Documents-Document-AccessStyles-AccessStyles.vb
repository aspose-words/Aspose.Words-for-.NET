' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
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
