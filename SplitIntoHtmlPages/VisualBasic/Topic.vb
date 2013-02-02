'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Namespace SplitIntoHtmlPages
	''' <summary>
	''' A simple class to hold a topic title and HTML file name together.
	''' </summary>
	Friend Class Topic
		Friend Sub New(ByVal title As String, ByVal fileName As String)
			mTitle = title
			mFileName = fileName
		End Sub

		Friend ReadOnly Property Title() As String
			Get
				Return mTitle
			End Get
		End Property

		Friend ReadOnly Property FileName() As String
			Get
				Return mFileName
			End Get
		End Property

		Private ReadOnly mTitle As String
		Private ReadOnly mFileName As String
	End Class
End Namespace
