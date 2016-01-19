Imports Microsoft.VisualBasic
Imports System
Imports System.Windows.Forms

Namespace DocumentExplorerExample

	''' <summary>
	''' Provides full information about application exception.
	''' </summary>
	Public Class ExceptionDialog
		Inherits Form
		#Region "Windows Form Designer generated code"

		Private components As System.ComponentModel.Container = Nothing
		Private buttonOk As System.Windows.Forms.Button
		Private text1 As System.Windows.Forms.TextBox

		Private Sub InitializeComponent()
			Dim resources As New System.Resources.ResourceManager(GetType(ExceptionDialog))
			Me.text1 = New System.Windows.Forms.TextBox()
			Me.buttonOk = New System.Windows.Forms.Button()
			Me.SuspendLayout()
			' 
			' text1
			' 
			Me.text1.Anchor = (CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) Or System.Windows.Forms.AnchorStyles.Left) Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.text1.AutoSize = False
			Me.text1.Location = New System.Drawing.Point(8, 8)
			Me.text1.Multiline = True
			Me.text1.Name = "text1"
			Me.text1.ReadOnly = True
			Me.text1.ScrollBars = System.Windows.Forms.ScrollBars.Both
			Me.text1.Size = New System.Drawing.Size(524, 244)
			Me.text1.TabIndex = 0
			Me.text1.Text = ""
			Me.text1.WordWrap = False
			' 
			' buttonOk
			' 
			Me.buttonOk.Anchor = (CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles))
			Me.buttonOk.DialogResult = System.Windows.Forms.DialogResult.OK
			Me.buttonOk.FlatStyle = System.Windows.Forms.FlatStyle.System
			Me.buttonOk.Location = New System.Drawing.Point(432, 260)
			Me.buttonOk.Name = "buttonOk"
			Me.buttonOk.Size = New System.Drawing.Size(100, 24)
			Me.buttonOk.TabIndex = 12
			Me.buttonOk.Text = "Continue"
			' 
			' ExceptionDialog
			' 
			Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
			Me.ClientSize = New System.Drawing.Size(540, 288)
			Me.Controls.Add(Me.buttonOk)
			Me.Controls.Add(Me.text1)
			Me.DockPadding.All = 8
			Me.Icon = (CType(resources.GetObject("$this.Icon"), System.Drawing.Icon))
			Me.MaximizeBox = False
			Me.MinimizeBox = False
			Me.Name = "ExceptionDialog"
			Me.ShowInTaskbar = False
			Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
			Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
			Me.Text = "Unexpected error occured"
			Me.ResumeLayout(False)

		End Sub

		Protected Overrides Overloads Sub Dispose(ByVal disposing As Boolean)
			If disposing Then
				If components IsNot Nothing Then
					components.Dispose()
				End If
			End If
			MyBase.Dispose(disposing)
		End Sub

		#End Region

		Public Sub New()
			InitializeComponent()
		End Sub

		Public Sub New(ByVal ex As Exception)
			InitializeComponent()
			Text = "Document Explorer - unexpected error occured"
			text1.Text = Constants.vbCrLf & Application.ProductName & ".exe " & Constants.vbCrLf & Constants.vbCrLf & "Version " & Application.ProductVersion & Constants.vbCrLf & Constants.vbCrLf & DateTime.Now.ToLongDateString() & " " & DateTime.Now.ToLongTimeString() & Constants.vbCrLf & Constants.vbCrLf & ex.ToString() & Constants.vbCrLf
			text1.SelectionStart = text1.Text.Length
		End Sub
	End Class
End Namespace