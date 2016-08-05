Imports System.IO
Imports System.Reflection
Imports Aspose.Words.Fonts
Imports Aspose.Words
Imports Aspose.Words.Saving
Public Class SetFontsFoldersSystemAndCustomFolder
    Public Shared Sub Run()
        ' ExStart:SetFontsFoldersSystemAndCustomFolder
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()
        Dim FontSettings As New FontSettings()

        Dim doc As New Document(dataDir & Convert.ToString("Rendering.doc"))

        ' Retrieve the array of environment-dependent font sources that are searched by default. For example this will contain a "Windows\Fonts\" source on a Windows machines.
        ' We add this array to a new ArrayList to make adding or removing font entries much easier.
        Dim fontSources As New ArrayList(FontSettings.GetFontsSources())

        ' Add a new folder source which will instruct Aspose.Words to search the following folder for fonts.
        Dim folderFontSource As New FolderFontSource("C:\MyFonts\", True)

        ' Add the custom folder which contains our fonts to the list of existing font sources.
        fontSources.Add(folderFontSource)

        ' Convert the Arraylist of source back into a primitive array of FontSource objects.
        Dim updatedFontSources As FontSourceBase() = DirectCast(fontSources.ToArray(GetType(FontSourceBase)), FontSourceBase())

        ' Apply the new set of font sources to use.
        FontSettings.SetFontsSources(updatedFontSources)
        ' Set font settings
        doc.FontSettings = FontSettings
        dataDir = dataDir & Convert.ToString("Rendering.SetFontsFolders_out_.pdf")
        doc.Save(dataDir)
        ' ExEnd:SetFontsFoldersSystemAndCustomFolder 
        Console.WriteLine(Convert.ToString(vbLf & "Fonts system and coustom folder is setup." & vbLf & "File saved at ") & dataDir)

    End Sub
End Class
