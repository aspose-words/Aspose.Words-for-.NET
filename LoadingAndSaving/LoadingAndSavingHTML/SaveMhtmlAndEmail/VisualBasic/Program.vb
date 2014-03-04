' Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection

Imports Aspose.Network.Mail
Imports Aspose.Words

Namespace SaveMhtmlAndEmail
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'ExStart
			'ExId:SaveMhtmlAndEmail
			'ExSummary:Shows how to save any document from Aspose.Words as MHTML and email using Aspose.Network.
			' Load the document into Aspose.Words.
			Dim srcFileName As String = Path.Combine(dataDir, "DinnerInvitationDemo.doc")
			Dim doc As New Document(srcFileName)

			' Save into a memory stream in MHTML format.
			Dim stream As Stream = New MemoryStream()
			doc.Save(stream, SaveFormat.Mhtml)
			' Rewind the stream to the beginning so Aspose.Network can read it.
			stream.Position = 0

			' Create an Aspose.Network MIME email message from the stream.
			Dim message As MailMessage = MailMessage.Load(stream, MessageFormat.Mht)
			message.From = "your_from@email.com"
			message.To = "your_to@email.com"
			message.Subject = "Aspose.Words + Aspose.Network MHTML Test Message"

			' Send the message using Aspose.Network
			Dim client As New SmtpClient()
			client.Host = "your_smtp.com"
			client.AuthenticationMethod = SmtpAuthentication.None
			client.Send(message)
			'ExEnd
		End Sub
	End Class
End Namespace