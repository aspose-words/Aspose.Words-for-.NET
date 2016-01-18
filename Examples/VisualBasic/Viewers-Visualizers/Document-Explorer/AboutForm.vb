Imports Microsoft.VisualBasic
Imports System
Namespace DocumentExplorerExample
	''' <summary>
	''' Shows an About form for this application.
	''' </summary>
	Public Class AboutForm
		Inherits System.Windows.Forms.Form

		#Region "Windows Form Designer generated code"

		Private components As System.ComponentModel.Container = Nothing
		Private pictureBox1 As System.Windows.Forms.PictureBox
		Private label1 As System.Windows.Forms.Label
		Private textBox1 As System.Windows.Forms.TextBox
		Private button1 As System.Windows.Forms.Button
		Private label2 As System.Windows.Forms.Label

		Private Sub InitializeComponent()
			Dim resources As New System.ComponentModel.ComponentResourceManager(GetType(AboutForm))
			Me.pictureBox1 = New System.Windows.Forms.PictureBox()
			Me.label1 = New System.Windows.Forms.Label()
			Me.label2 = New System.Windows.Forms.Label()
			Me.textBox1 = New System.Windows.Forms.TextBox()
			Me.button1 = New System.Windows.Forms.Button()
			Me.SuspendLayout()
			' 
			' pictureBox1
			' 
			Me.pictureBox1.BackColor = System.Drawing.Color.Transparent
			Me.pictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
			Me.pictureBox1.Image = (CType(resources.GetObject("pictureBox1.Image"), System.Drawing.Image))
			Me.pictureBox1.Location = New System.Drawing.Point(8, 8)
			Me.pictureBox1.Name = "pictureBox1"
			Me.pictureBox1.Size = New System.Drawing.Size(128, 132)
			Me.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
			Me.pictureBox1.TabIndex = 0
			Me.pictureBox1.TabStop = False
			' 
			' label1
			' 
			Me.label1.Font = New System.Drawing.Font("Times New Roman", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, (CByte(204)))
			Me.label1.ForeColor = System.Drawing.Color.Black
			Me.label1.Location = New System.Drawing.Point(144, 4)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(404, 32)
			Me.label1.TabIndex = 1
			Me.label1.Text = "Document Explorer Demo for Aspose.Words "
			Me.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
			' 
			' label2
			' 
			Me.label2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
			Me.label2.Font = New System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, (CByte(204)))
			Me.label2.ForeColor = System.Drawing.Color.Black
			Me.label2.Location = New System.Drawing.Point(144, 120)
			Me.label2.Name = "label2"
			Me.label2.Size = New System.Drawing.Size(364, 20)
			Me.label2.TabIndex = 2
			Me.label2.Text = "Copyright © 2002-2010 Aspose Pty Ltd. All Rights Reserved. "
			Me.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
			' 
			' textBox1
			' 
			Me.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
			Me.textBox1.Location = New System.Drawing.Point(144, 40)
			Me.textBox1.Multiline = True
			Me.textBox1.Name = "textBox1"
			Me.textBox1.ReadOnly = True
			Me.textBox1.Size = New System.Drawing.Size(428, 68)
			Me.textBox1.TabIndex = 3
			Me.textBox1.TabStop = False
			Me.textBox1.Text = resources.GetString("textBox1.Text")
			' 
			' button1
			' 
			Me.button1.DialogResult = System.Windows.Forms.DialogResult.OK
			Me.button1.ForeColor = System.Drawing.SystemColors.ControlText
			Me.button1.Location = New System.Drawing.Point(508, 120)
			Me.button1.Name = "button1"
			Me.button1.Size = New System.Drawing.Size(75, 23)
			Me.button1.TabIndex = 4
			Me.button1.Text = "OK"
			' 
			' AboutForm
			' 
			Me.AcceptButton = Me.button1
			Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
			Me.BackColor = System.Drawing.SystemColors.Control
			Me.ClientSize = New System.Drawing.Size(590, 150)
			Me.ControlBox = False
			Me.Controls.Add(Me.button1)
			Me.Controls.Add(Me.textBox1)
			Me.Controls.Add(Me.label2)
			Me.Controls.Add(Me.label1)
			Me.Controls.Add(Me.pictureBox1)
			Me.ForeColor = System.Drawing.SystemColors.Window
			Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
			Me.Name = "AboutForm"
			Me.ShowInTaskbar = False
			Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
			Me.TopMost = True
			Me.ResumeLayout(False)
			Me.PerformLayout()

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
	End Class
End Namespace