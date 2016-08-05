Imports System.IO
Imports System.Reflection
Imports Aspose.Words.Fonts
Imports Aspose.Words
Imports Aspose.Words.Saving
Public Class ReceiveNotificationsOfFont
    Public Shared Sub Run()
        ' ExStart:ReceiveNotificationsOfFonts 
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_RenderingAndPrinting()
        Dim FontSettings As New FontSettings()


        Dim doc As New Document(dataDir & Convert.ToString("Rendering.doc"))
        ' We can choose the default font to use in the case of any missing fonts.
        FontSettings.DefaultFontName = "Arial"
        ' For testing we will set Aspose.Words to look for fonts only in a folder which doesn't exist. Since Aspose.Words won't
        ' find any fonts in the specified directory, then during rendering the fonts in the document will be subsuited with the default
        ' font specified under FontSettings.DefaultFontName. We can pick up on this subsuition using our callback.
        FontSettings.SetFontsFolder(String.Empty, False)

        ' Create a new class implementing IWarningCallback which collect any warnings produced during document save.
        Dim callback As New HandleDocumentWarnings()

        doc.WarningCallback = callback
        ' Set font settings
        doc.FontSettings = FontSettings
        Dim path As String = dataDir & Convert.ToString("Rendering.MissingFontNotification_out_.pdf")
        ' Pass the save options along with the save path to the save method.
        doc.Save(path)
        ' ExEnd:ReceiveNotificationsOfFonts 
        Console.WriteLine(Convert.ToString(vbLf & "Receive notifications of font substitutions by using IWarningCallback processed." & vbLf & "File saved at ") & path)

        ReceiveWarningNotification(doc, dataDir)
    End Sub
    Private Shared Sub ReceiveWarningNotification(doc As Document, dataDir As String)
        ' ExStart:ReceiveWarningNotification 
        ' When you call UpdatePageLayout the document is rendered in memory. Any warnings that occured during rendering
        ' are stored until the document save and then sent to the appropriate WarningCallback.
        doc.UpdatePageLayout()

        ' Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class.
        Dim callback As New HandleDocumentWarnings()

        doc.WarningCallback = callback
        dataDir = dataDir & Convert.ToString("Rendering.FontsNotificationUpdatePageLayout_out_.pdf")
        ' Even though the document was rendered previously, any save warnings are notified to the user during document save.
        doc.Save(dataDir)
        ' ExEnd:ReceiveWarningNotification  
    End Sub
    ' ExStart:HandleDocumentWarnings
    Public Class HandleDocumentWarnings
        Implements IWarningCallback
        ''' <summary>
        ''' Our callback only needs to implement the "Warning" method. This method is called whenever there is a
        ''' potential issue during document procssing. The callback can be set to listen for warnings generated during document
        ''' load and/or document save.
        ''' </summary>
        Public Sub Warning(ByVal info As WarningInfo) Implements IWarningCallback.Warning
            ' We are only interested in fonts being substituted.
            If info.WarningType = WarningType.FontSubstitution Then
                Console.WriteLine("Font substitution: " & info.Description)
            End If
        End Sub

    End Class
    ' ExEnd:HandleDocumentWarnings
End Class
