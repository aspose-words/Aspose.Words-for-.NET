'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Net
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Documents
Imports System.Windows.Ink
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Animation
Imports System.Windows.Shapes

Namespace FirstFloor.Documents.Aspose
	Public Class DocumentInfo
		Private privateName As String
		Public Property Name() As String
			Get
				Return privateName
			End Get
			Set(ByVal value As String)
				privateName = value
			End Set
		End Property
		Private privateDescription As String
		Public Property Description() As String
			Get
				Return privateDescription
			End Get
			Set(ByVal value As String)
				privateDescription = value
			End Set
		End Property
		Private privateXpsLocation As Uri
		Public Property XpsLocation() As Uri
			Get
				Return privateXpsLocation
			End Get
			Set(ByVal value As Uri)
				privateXpsLocation = value
			End Set
		End Property
		Private privateOriginalLocation As Uri
		Public Property OriginalLocation() As Uri
			Get
				Return privateOriginalLocation
			End Get
			Set(ByVal value As Uri)
				privateOriginalLocation = value
			End Set
		End Property
	End Class
End Namespace
