'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.Text.RegularExpressions

Namespace Word2HelpExample
	''' <summary>
	''' Central storage for regular expressions used in the project.
	''' </summary>
	Public Class RegularExpressions
		' This class is static. No instance creation is allowed.
		Private Sub New()
		End Sub

		''' <summary>
		''' Regular expression specifying html title (framing tags excluded).
		''' </summary>
		Public Shared ReadOnly Property HtmlTitle() As Regex
			Get
				If gHtmlTitle Is Nothing Then
					gHtmlTitle = New Regex(HtmlTitlePattern, RegexOptions.IgnoreCase Or RegexOptions.Singleline Or RegexOptions.Compiled)
				End If
				Return gHtmlTitle
			End Get
		End Property

		''' <summary>
		''' Regular expression specifying html head.
		''' </summary>
		Public Shared ReadOnly Property HtmlHead() As Regex
			Get
				If gHtmlHead Is Nothing Then
					gHtmlHead = New Regex(HtmlHeadPattern, RegexOptions.IgnoreCase Or RegexOptions.Singleline Or RegexOptions.Compiled)
				End If
				Return gHtmlHead
			End Get
		End Property

		''' <summary>
		''' Regular expression specifying space right after div keyword in the first div declaration of html body.
		''' </summary>
		Public Shared ReadOnly Property HtmlBodyDivStart() As Regex
			Get
				If gHtmlBodyDivStart Is Nothing Then
					gHtmlBodyDivStart = New Regex(HtmlBodyDivStartPattern, RegexOptions.IgnoreCase Or RegexOptions.Singleline Or RegexOptions.Compiled)
				End If
				Return gHtmlBodyDivStart
			End Get
		End Property

		Private Const HtmlTitlePattern As String = "(?<=\<title\>).*?(?=\</title\>)"
		Private Shared gHtmlTitle As Regex

		Private Const HtmlHeadPattern As String = "\<head\>.*?\</head\>"
		Private Shared gHtmlHead As Regex

		Private Const HtmlBodyDivStartPattern As String = "(?<=\<body\>\s*\<div)\s"
		Private Shared gHtmlBodyDivStart As Regex
	End Class
End Namespace