Imports Aspose.Words
Imports Aspose.Words.Loading
Imports System.Net
' ExStart:ImageLoadingWithCredentialsHandler
Public Class ImageLoadingWithCredentialsHandler

    Public Sub New()
        mWebClient = New WebClient()
    End Sub
    Public Function ResourceLoading(ByVal args As ResourceLoadingArgs) As ResourceLoadingAction
        If args.ResourceType = ResourceType.Image Then
            Dim uri As Uri = New Uri(args.Uri)

            If uri.Host = "www.aspose.com" Then
                mWebClient.Credentials = New NetworkCredential("User1", "akjdlsfkjs")
            Else
                mWebClient.Credentials = New NetworkCredential("SomeOtherUserID", "wiurlnlvs")
            End If

            ' Download the bytes from the location referenced by the URI.
            Dim imageBytes As Byte() = mWebClient.DownloadData(args.Uri)

            args.SetData(imageBytes)

            Return ResourceLoadingAction.UserProvided
        Else
            Return ResourceLoadingAction.Default
        End If
    End Function

    Private mWebClient As WebClient
End Class
' ExEnd:ImageLoadingWithCredentialsHandler
