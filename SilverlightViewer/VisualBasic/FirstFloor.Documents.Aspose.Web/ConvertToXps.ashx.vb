'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Web

Imports Aspose.Words

Namespace FirstFloor.Documents.Aspose.Web
	Public Class ConvertToXps
		Implements IHttpHandler
		Public Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest
			If context.Request.ContentLength > 2 << 18 Then
				Throw New NotSupportedException()
			End If
			Dim document = New Document(context.Request.InputStream)
			document.Save(context.Response.OutputStream, SaveFormat.Xps)
		End Sub

		Public ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
			Get
				Return False
			End Get
		End Property
	End Class
End Namespace
