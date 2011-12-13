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
	Public Class DocumentSelectedEventArgs
		Inherits EventArgs
		Public Sub New(ByVal document As DocumentInfo)
			Me.Document = document
		End Sub

		Private privateDocument As DocumentInfo
		Public Property Document() As DocumentInfo
			Get
				Return privateDocument
			End Get
			Private Set(ByVal value As DocumentInfo)
				privateDocument = value
			End Set
		End Property
	End Class
End Namespace
