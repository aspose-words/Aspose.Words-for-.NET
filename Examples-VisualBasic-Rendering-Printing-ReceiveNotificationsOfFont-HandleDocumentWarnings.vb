' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
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
