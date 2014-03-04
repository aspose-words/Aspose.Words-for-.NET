' Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.

Imports Microsoft.VisualBasic
Imports System
Namespace Excel2WordExample
	Partial Public Class MainForm
		''' <summary>
		''' Required designer variable.
		''' </summary>
		Private components As System.ComponentModel.IContainer = Nothing

		''' <summary>
		''' Clean up any resources being used.
		''' </summary>
		''' <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		Protected Overrides Sub Dispose(ByVal disposing As Boolean)
			If disposing AndAlso (components IsNot Nothing) Then
				components.Dispose()
			End If
			MyBase.Dispose(disposing)
		End Sub

		#Region "Windows Form Designer generated code"

		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
			Dim resources As New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
			Me.labelStatus = New System.Windows.Forms.Label()
			Me.buttonConvert = New System.Windows.Forms.Button()
			Me.openFileDialog = New System.Windows.Forms.OpenFileDialog()
			Me.saveFileDialog = New System.Windows.Forms.SaveFileDialog()
			Me.SuspendLayout()
			' 
			' labelStatus
			' 
			Me.labelStatus.AutoSize = True
			Me.labelStatus.Location = New System.Drawing.Point(120, 17)
			Me.labelStatus.Name = "labelStatus"
			Me.labelStatus.Size = New System.Drawing.Size(0, 13)
			Me.labelStatus.TabIndex = 1
			' 
			' buttonConvert
			' 
			Me.buttonConvert.Anchor = System.Windows.Forms.AnchorStyles.None
			Me.buttonConvert.Location = New System.Drawing.Point(74, 39)
			Me.buttonConvert.Name = "buttonConvert"
			Me.buttonConvert.Size = New System.Drawing.Size(138, 26)
			Me.buttonConvert.TabIndex = 3
			Me.buttonConvert.Text = "Convert..."
			Me.buttonConvert.UseVisualStyleBackColor = True
'			Me.buttonConvert.Click += New System.EventHandler(Me.buttonConvert_Click);
			' 
			' openFileDialog
			' 
			Me.openFileDialog.Filter = "Excel files (*.xls, *.xlsx)|*.xls;*.xlsx"
			' 
			' saveFileDialog
			' 
			Me.saveFileDialog.Filter = resources.GetString("saveFileDialog.Filter")
			' 
			' MainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(287, 105)
			Me.Controls.Add(Me.buttonConvert)
			Me.Controls.Add(Me.labelStatus)
			Me.Name = "MainForm"
			Me.Text = "MainForm"
			Me.ResumeLayout(False)
			Me.PerformLayout()

		End Sub

		#End Region

		Private labelStatus As System.Windows.Forms.Label
		Private WithEvents buttonConvert As System.Windows.Forms.Button
		Private openFileDialog As System.Windows.Forms.OpenFileDialog
		Private saveFileDialog As System.Windows.Forms.SaveFileDialog
	End Class
End Namespace