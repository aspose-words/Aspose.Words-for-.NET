'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Navigation

Namespace FirstFloor.Documents.Aspose
	Partial Public Class MainPage
		Inherits UserControl
		Public Sub New()
			InitializeComponent()
		End Sub

		' After the Frame navigates, ensure the HyperlinkButton representing the current page is selected
		Private Sub ContentFrame_Navigated(ByVal sender As Object, ByVal e As NavigationEventArgs)
			For Each child As UIElement In LinksStackPanel.Children
				Dim hb As HyperlinkButton = TryCast(child, HyperlinkButton)
				If hb IsNot Nothing AndAlso hb.NavigateUri IsNot Nothing Then
					If hb.NavigateUri.ToString().Equals(e.Uri.ToString()) Then
						VisualStateManager.GoToState(hb, "ActiveLink", True)
					Else
						VisualStateManager.GoToState(hb, "InactiveLink", True)
					End If
				End If
			Next child
		End Sub

		' If an error occurs during navigation, show an error window
		Private Sub ContentFrame_NavigationFailed(ByVal sender As Object, ByVal e As NavigationFailedEventArgs)
			e.Handled = True
			Dim errorWin As ChildWindow = New ErrorWindow(e.Uri)
			errorWin.Show()
		End Sub
	End Class
End Namespace