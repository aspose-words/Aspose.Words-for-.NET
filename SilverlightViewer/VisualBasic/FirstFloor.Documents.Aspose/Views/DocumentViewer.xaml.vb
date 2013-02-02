'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Windows
Imports System.Windows.Browser
Imports System.Windows.Controls

Imports FirstFloor.Documents.Controls
Imports FirstFloor.Documents.Extensions
Imports FirstFloor.Documents.IO

Namespace FirstFloor.Documents.Aspose.Views
	Partial Public Class DocumentViewer
		Inherits UserControl
		Private Class UploadRequest
			Private privateFileName As String
			Public Property FileName() As String
				Get
					Return privateFileName
				End Get
				Set(ByVal value As String)
					privateFileName = value
				End Set
			End Property
			Private privateFileStream As Stream
			Public Property FileStream() As Stream
				Get
					Return privateFileStream
				End Get
				Set(ByVal value As Stream)
					privateFileStream = value
				End Set
			End Property
			Private privateRequest As HttpWebRequest
			Public Property Request() As HttpWebRequest
				Get
					Return privateRequest
				End Get
				Set(ByVal value As HttpWebRequest)
					privateRequest = value
				End Set
			End Property
		End Class

		Public Event Close As EventHandler
		Private client As XpsClient

		Public Sub New()
			InitializeComponent()

			Me.client = New XpsClient()
			AddHandler Me.client.LoadXpsDocumentCompleted, AddressOf client_LoadXpsDocumentCompleted

			Me.toolbar.Viewer = Me.viewer
		End Sub

		Public ReadOnly Property Viewer() As FixedDocumentViewer
			Get
				Return Me.viewer
			End Get
		End Property

		Private Sub client_LoadXpsDocumentCompleted(ByVal sender As Object, ByVal e As LoadXpsDocumentCompletedEventArgs)
			If (Not e.Cancelled) Then
				If e.Error IsNot Nothing Then
					ErrorWindow.ShowError(e.Error)
				Else
					Me.viewer.FixedDocument = e.Document.FixedDocuments.FirstOrDefault()
					Me.toolbar.Refresh()
				End If
			End If
		End Sub

		Public Sub ClearDocument()
			If Me.viewer.FixedDocument IsNot Nothing Then
				Me.viewer.FixedDocument.Owner.Dispose()
			End If
			Me.viewer.FixedDocument = Nothing
			Me.title.Text = Nothing
			Me.toolbar.OriginalUri = Nothing
			Me.toolbar.Refresh()
		End Sub

		Public Sub LoadDocument(ByVal title As String, ByVal uri As Uri, ByVal originalUri As Uri)
			ClearDocument()

			Me.title.Text = title
			Me.toolbar.OriginalUri = originalUri
			Me.loading.Visibility = Visibility.Visible
			Me.status.Text = "Loading...Please Wait"

			Dim webClient = New WebClient()
			AddHandler webClient.OpenReadCompleted, Function(o, e)
				If (Not e.Cancelled) Then
					If e.Error IsNot Nothing Then
						ErrorWindow.ShowError(e.Error)
					Else
						LoadDocument(e.Result)
					End If
				End If
				Me.loading.Visibility = Visibility.Collapsed

			webClient.OpenReadAsync(uri)
		End Sub

		Public Sub LoadLocalDocument(ByVal file As FileInfo)
			Try
				Dim uri = New Uri(HtmlPage.Document.DocumentUri, "ConvertToXps.ashx")
				Dim request = CType(WebRequest.Create(uri), HttpWebRequest)
				Dim uploadRequest = New UploadRequest() With {.FileName = file.Name, .FileStream = file.OpenRead(), .Request = request}
				request.Method = "POST"
				request.BeginGetRequestStream(AddressOf OnGetRequestStream, uploadRequest)

				ClearDocument()
				Me.loading.Visibility = Visibility.Visible
				Me.status.Text = "Sending document...Please Wait"
				Me.title.Text = uploadRequest.FileName
			Catch ex As Exception
				ErrorWindow.ShowError(ex)
				Me.loading.Visibility = Visibility.Collapsed
			End Try
		End Sub

		Private Sub LoadDocument(ByVal stream As Stream)
			Dim reader = New SharpZipPackageReader(stream)

			Dim settings = New LoadXpsDocumentSettings() With {.IncludeProperties = False, .IncludeDocumentStructures = False, .IncludeAnnotations = False}

			Me.client.LoadXpsDocumentAsync(reader, settings)
		End Sub

		Private Sub OnGetRequestStream(ByVal result As IAsyncResult)
			Dim uploadRequest = CType(result.AsyncState, UploadRequest)
			Try
				Using stream = uploadRequest.Request.EndGetRequestStream(result)
					Dim buffer = New Byte(4095){}
					Dim bytesRead As Integer

					bytesRead = uploadRequest.FileStream.Read(buffer, 0, buffer.Length)
					Do While bytesRead <> 0
						stream.Write(buffer, 0, bytesRead)
						bytesRead = uploadRequest.FileStream.Read(buffer, 0, buffer.Length)
					Loop
				End Using

'TODO: INSTANT VB TODO TASK: Assignments within expressions are not supported in VB.NET
'ORIGINAL LINE: Dispatcher.BeginInvoke(() => { Me.status.Text = "Receiving XPS...Please Wait";
				Dispatcher.BeginInvoke(Function() { Me.status.Text = "Receiving XPS...Please Wait"
			End Try
			   )

				uploadRequest.Request.BeginGetResponse(OnGetResponse, uploadRequest)
		End Sub
			Private Function [catch](ByVal e As Exception) As [Private]
'TODO: INSTANT VB TODO TASK: Assignments within expressions are not supported in VB.NET
'ORIGINAL LINE: Dispatcher.BeginInvoke(() => { Me.loading.Visibility = Visibility.Collapsed; ErrorWindow.ShowError(e);
				Dispatcher.BeginInvoke(Function() { Me.loading.Visibility = Visibility.Collapsed; ErrorWindow.ShowError(e)
			End Function
			   Private )
	End Class
			Finally
				RaiseEvent uploadRequest.FileStream.Close()
			End Try
End Namespace

		private void OnGetResponse(IAsyncResult result)
			Try
				Dim uploadRequest = CType(result.AsyncState, UploadRequest)
				Dim response = CType(uploadRequest.Request.EndGetResponse(result), HttpWebResponse)

'TODO: INSTANT VB TODO TASK: Assignments within expressions are not supported in VB.NET
'ORIGINAL LINE: Dispatcher.BeginInvoke(() => { LoadDocument(response.GetResponseStream()); Me.loading.Visibility = Visibility.Collapsed;
				Dispatcher.BeginInvoke(Function() { LoadDocument(response.GetResponseStream()); Me.loading.Visibility = Visibility.Collapsed
			End Try
			   )

			Catch e As Exception
'TODO: INSTANT VB TODO TASK: Assignments within expressions are not supported in VB.NET
'ORIGINAL LINE: Dispatcher.BeginInvoke(() => { Me.loading.Visibility = Visibility.Collapsed; ErrorWindow.ShowError(e);
				Dispatcher.BeginInvoke(Function() { Me.loading.Visibility = Visibility.Collapsed; ErrorWindow.ShowError(e)
			End Try
			   )
			}
		}

		private void close_Click(Object sender, RoutedEventArgs e)
			Application.Current.Host.Content.IsFullScreen = False
			RaiseEvent Close(Me, EventArgs.Empty)

		private void viewer_LinkClick(Object sender, LinkClickEventArgs e)
			If e.NavigateUri.IsAbsoluteUri AndAlso (e.NavigateUri.Scheme = Uri.UriSchemeHttp OrElse e.NavigateUri.Scheme = Uri.UriSchemeHttps) Then
				HtmlPage.Window.Navigate(e.NavigateUri, "_blank")
			End If
	}
}
