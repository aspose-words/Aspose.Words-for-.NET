' Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Windows.Forms

Namespace DocumentExplorer
	''' <summary>
	''' Lets the user specify the page number to go to.
	''' </summary>
	Public Class GoToPageForm
		Inherits System.Windows.Forms.Form
		Private components As System.ComponentModel.Container = Nothing
		Private promptLabel As System.Windows.Forms.Label
		Private WithEvents pageNumberTextBox As System.Windows.Forms.TextBox
		Private mPageNumber As Integer
		Private WithEvents okBtn As System.Windows.Forms.Button
		Private cancelBtn As System.Windows.Forms.Button
		Private mMaxPageNumber As Integer = 1

		Public Sub New()
			InitializeComponent()
		End Sub

		Public WriteOnly Property MaxPageNumber() As Integer
			Set(ByVal value As Integer)
				mMaxPageNumber = value
			End Set
		End Property

		Public ReadOnly Property PageNumber() As Integer
			Get
				Return mPageNumber
			End Get
		End Property

		''' <summary>
		''' Clean up any resources being used.
		''' </summary>
		Protected Overrides Overloads Sub Dispose(ByVal disposing As Boolean)
			If disposing Then
				If components IsNot Nothing Then
					components.Dispose()
				End If
			End If
			MyBase.Dispose(disposing)
		End Sub

		#Region "Windows Form Designer generated code"
		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
			Me.promptLabel = New System.Windows.Forms.Label()
			Me.pageNumberTextBox = New System.Windows.Forms.TextBox()
			Me.okBtn = New System.Windows.Forms.Button()
			Me.cancelBtn = New System.Windows.Forms.Button()
			Me.SuspendLayout()
			' 
			' promptLabel
			' 
			Me.promptLabel.Location = New System.Drawing.Point(16, 16)
			Me.promptLabel.Name = "promptLabel"
			Me.promptLabel.Size = New System.Drawing.Size(192, 23)
			Me.promptLabel.TabIndex = 0
			' 
			' pageNumberTextBox
			' 
			Me.pageNumberTextBox.Location = New System.Drawing.Point(16, 40)
			Me.pageNumberTextBox.MaxLength = 5
			Me.pageNumberTextBox.Name = "pageNumberTextBox"
			Me.pageNumberTextBox.Size = New System.Drawing.Size(192, 20)
			Me.pageNumberTextBox.TabIndex = 1
			Me.pageNumberTextBox.Text = ""
'			Me.pageNumberTextBox.TextChanged += New System.EventHandler(Me.pageNumberTextBox_TextChanged);
			' 
			' okBtn
			' 
			Me.okBtn.DialogResult = System.Windows.Forms.DialogResult.OK
			Me.okBtn.Enabled = False
			Me.okBtn.Location = New System.Drawing.Point(56, 72)
			Me.okBtn.Name = "okBtn"
			Me.okBtn.TabIndex = 2
			Me.okBtn.Text = "OK"
'			Me.okBtn.Click += New System.EventHandler(Me.okButton_Click);
			' 
			' cancelBtn
			' 
			Me.cancelBtn.CausesValidation = False
			Me.cancelBtn.DialogResult = System.Windows.Forms.DialogResult.Cancel
			Me.cancelBtn.Location = New System.Drawing.Point(136, 72)
			Me.cancelBtn.Name = "cancelBtn"
			Me.cancelBtn.TabIndex = 3
			Me.cancelBtn.Text = "Cancel"
			' 
			' GoToPageForm
			' 
			Me.AcceptButton = Me.okBtn
			Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
			Me.CancelButton = Me.cancelBtn
			Me.CausesValidation = False
			Me.ClientSize = New System.Drawing.Size(218, 104)
			Me.Controls.Add(Me.okBtn)
			Me.Controls.Add(Me.pageNumberTextBox)
			Me.Controls.Add(Me.promptLabel)
			Me.Controls.Add(Me.cancelBtn)
			Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
			Me.MaximizeBox = False
			Me.MinimizeBox = False
			Me.Name = "GoToPageForm"
			Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
			Me.Text = "Go to Page"
'			Me.Load += New System.EventHandler(Me.GoToPageForm_Load);
			Me.ResumeLayout(False)

		End Sub
		#End Region

		Private Sub GoToPageForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
			promptLabel.Text = String.Format("Enter page number (1-{0})", mMaxPageNumber)
		End Sub

		Private Sub pageNumberTextBox_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles pageNumberTextBox.TextChanged
			okBtn.Enabled = (pageNumberTextBox.Text.Length > 0)
		End Sub

		Private Sub okButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles okBtn.Click
			Try
				If (Not TryParse(pageNumberTextBox.Text, mPageNumber)) Then
					Throw New Exception("Please enter a valid page number.")
				End If

				If (mPageNumber < 1) OrElse (mPageNumber > mMaxPageNumber) Then
					Throw New Exception(String.Format("Page number must be between 1 and {0}.", mMaxPageNumber))
				End If
			Catch ex As Exception
				MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
				DialogResult = System.Windows.Forms.DialogResult.None
			End Try
		End Sub

		Private Shared Function TryParse(ByVal text As String, <System.Runtime.InteropServices.Out()> ByRef value As Integer) As Boolean
			value = 0
			Try
				value = Integer.Parse(text)
				Return True
			Catch e1 As Exception
				Return False
			End Try
		End Function
	End Class
End Namespace
