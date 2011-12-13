'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Namespace AjaxGenerateDoc
	''' <summary>
	''' Demonstrates how to show a progress message while invoking Aspose.Words to generate a document. 
	''' See the .aspx file for more code.
	''' </summary>
	Partial Public Class ExampleUsingIFrame1
		Inherits System.Web.UI.Page
		''' <summary>
		''' Check generation of a document is complete or not.
		''' Called from the script on the web page.
		''' </summary>
		<System.Web.Services.WebMethod(EnableSession := True)> _
		Public Shared Function CheckCompleted() As Boolean
			Dim completed As Boolean = False
			If System.Web.HttpContext.Current.Session("Completed") IsNot Nothing Then
				'Get a value from a Session variable that was set by GenerateFile.aspx.
				completed = System.Convert.ToBoolean(System.Web.HttpContext.Current.Session("Completed"))
			End If
			Return completed
		End Function
	End Class
End Namespace

