Imports Aspose.Words
Imports Aspose.Words.Saving
Imports System.IO

Public Class SpecifySaveOption
    Public Shared Sub Run()
        ' ExStart:SpecifySaveOption            
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()

        Dim fileName As String = "TestFile RenderShape.docx"

        Dim doc As New Document(dataDir & fileName)

        ' This is the directory we want the exported images to be saved to.
        Dim imagesDir As String = Path.Combine(dataDir, "Images")

        ' The folder specified needs to exist and should be empty.
        If Directory.Exists(imagesDir) Then
            Directory.Delete(imagesDir, True)
        End If

        Directory.CreateDirectory(imagesDir)

        ' Set an option to export form fields as plain text, not as HTML input elements.
        Dim options As New HtmlSaveOptions(SaveFormat.Html)
        options.ExportTextInputFormFieldAsText = True
        options.ImagesFolder = imagesDir

        dataDir = dataDir & Convert.ToString("Document.SaveWithOptions_out.html")
        doc.Save(dataDir, options)

        ' ExEnd:SpecifySaveOption

        Console.WriteLine(Convert.ToString(vbLf & "Save option specified successfully." & vbLf & "File saved at ") & dataDir)

    End Sub
End Class
