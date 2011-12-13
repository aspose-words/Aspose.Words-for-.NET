'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Specialized
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Text
Imports System.Windows.Forms
Imports Aspose.Words
Imports Aspose.Words.Saving
Imports System.Web.UI.WebControls

Namespace SaveHtmlAndEmail
	Partial Public Class MainForm
		Inherits Form
		Private inputFileName As String = String.Empty

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub buttonOpenDocument_Click(ByVal sender As Object, ByVal e As EventArgs) Handles buttonOpenDocument.Click
			labelMessage.Visible = False

			Try
				' Prompt the user to choose the input document.
				If openDocumentFileDialog.ShowDialog().Equals(System.Windows.Forms.DialogResult.OK) Then
					inputFileName = openDocumentFileDialog.FileName
					panelSend.Enabled = True
				End If
			Catch ex As Exception
				MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error)
			End Try
		End Sub

		Private Sub buttonSend_Click(ByVal sender As Object, ByVal e As EventArgs) Handles buttonSend.Click
			' Extract information that is needed to send email message from the user interface.
			Dim smtp As String = textboxSmtp.Text ' Your smtp server
			Dim emailFrom As String = textboxEmailFrom.Text ' Your email
			Dim password As String = textboxPassword.Text ' Your password
			Dim emailTo As String = textboxEmailTo.Text ' Recipient email
			Dim subject As String = textboxSubject.Text ' Subject
			Dim useAuth As Boolean = checkboxAuth.Checked ' Use authentication

			Dim port As Integer ' The port to use
			Integer.TryParse(textboxPort.Text, port)

			If port = 0 Then
				port = 25 ' If the port was not defined it will be parsed as 0. Change to default port 25.
			End If

			Try
				labelMessage.Visible = False
				buttonSend.Enabled = False

				' Send the information required to send the e-mail.
				Send(smtp, emailFrom, password, emailTo, subject, port, useAuth, inputFileName)

				' Show message
				labelMessage.Visible = True
			Catch ex As Exception
				MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error)

			Finally
				' Even if an exception occurs reset the send button.
				buttonSend.Enabled = True
			End Try
		End Sub

		''' <summary>
		''' Convert document to HTML mail message and send it to recipient
		''' </summary>
		''' <param name="smtp">Smtp server</param>
		''' <param name="emailFrom">Sender e-mail</param>
		''' <param name="password">Sender password</param>
		''' <param name="emailTo">Recipient e-mail</param>
		''' <param name="subject">E-mail subject</param>
		''' <param name="port">Port to use</param>
		''' <param name="useAuth">Specify authentication</param>
		''' <param name="inputFileName">Document file name</param>
		Private Shared Sub Send(ByVal smtp As String, ByVal emailFrom As String, ByVal password As String, ByVal emailTo As String, ByVal subject As String, ByVal port As Integer, ByVal useAuth As Boolean, ByVal inputFileName As String)
			' Create temporary folder for Aspose.Words to store images to during export.
			Dim tempDir As String = Path.Combine(Path.GetTempPath(), "AsposeMail")
			If (Not Directory.Exists(tempDir)) Then
				Directory.CreateDirectory(tempDir)
			End If

			' Open the document.
			Dim doc As New Document(inputFileName)

			Dim saveOptions As New HtmlSaveOptions()
			' Specify folder where images will be saved.
			saveOptions.ImagesFolder = tempDir
			' We want the images in the HTML to be referenced in the e-mail as attachments so add the cid prefix to the image file name.
			' This replaces what would be the path to the image with the "cid" prefix.
			saveOptions.ImagesFolderAlias = "cid:"
			' Header footers don't normally export well in HTML format so remove them.
			saveOptions.ExportHeadersFootersMode = ExportHeadersFootersMode.None

			' Save the document to stream in HTML format.
			Dim htmlStream As New MemoryStream()
			doc.Save(htmlStream, saveOptions)

			' Read the HTML from the stream as plain text.
			Dim htmlText As String = Encoding.UTF8.GetString(htmlStream.ToArray())
			htmlStream.Close()

			' Save the HTML into the temporary folder.
			Dim htmlFile As Stream = New FileStream(Path.Combine(tempDir, "Message.html"), FileMode.Create)
			Dim htmlWriter As New StreamWriter(htmlFile)
			htmlWriter.Write(htmlText)
			htmlWriter.Close()
			htmlFile.Close()

			' Create the mail definiton and specify the appropriate settings.
			Dim mail As New MailDefinition()
			mail.IsBodyHtml = True
			mail.BodyFileName = Path.Combine(tempDir, "Message.html")
			mail.From = emailFrom
			mail.Subject = subject

			' Get the names of the images in the temporary folder.
			Dim fileNames() As String = Directory.GetFiles(tempDir)

			' Add each image as an embedded object to the message.
			For imgIndex As Integer = 0 To fileNames.Length - 1
				Dim imgFullName As String = fileNames(imgIndex)
				Dim imgName As String = Path.GetFileName(fileNames(imgIndex))
				' The ID of the embedded object is the name of the image preceeded with a foward slash.
				mail.EmbeddedObjects.Add(New EmbeddedMailObject(String.Format("/{0}", imgName), imgFullName))
			Next imgIndex

			Dim message As MailMessage = Nothing

			' Create the message.
			Try
				message = mail.CreateMailMessage(emailTo, New ListDictionary(), New System.Web.UI.Control())

				' Create the SMTP client to send the message with.
				Dim sender As New SmtpClient(smtp)

				' Set the credentials.
				sender.Credentials = New NetworkCredential(emailFrom, password)
				' Set port.
				sender.Port = port
				' Choose to enable authentication.
				sender.EnableSsl = useAuth

				' Send the e-mail message.
				sender.Send(message)

			Catch e As Exception
				Throw e

			Finally
				' This frees the Message.html file if an exception occurs.
				message.Dispose()
			End Try

			' Delete the temp folder.
			Directory.Delete(tempDir, True)
		End Sub

		''' <summary>
		''' This restricts the user entering anything but digits up to a certain length in the port textbox.
		''' </summary>
		Private Sub textBoxPort_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles textboxPort.KeyPress
			If ((Not Char.IsControl(e.KeyChar)) AndAlso (Not Char.IsDigit(e.KeyChar))) OrElse ((Not Char.IsControl(e.KeyChar)) AndAlso (CType(sender, System.Windows.Forms.TextBox)).Text.Length >= 5) Then
				e.Handled = True
			End If
		End Sub
	End Class
End Namespace