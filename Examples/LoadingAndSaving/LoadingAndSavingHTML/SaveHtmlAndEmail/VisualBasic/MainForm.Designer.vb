' Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Namespace SaveHtmlAndEmailExample
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
			Me.buttonOpenDocument = New System.Windows.Forms.Button()
			Me.textboxSmtp = New System.Windows.Forms.TextBox()
			Me.textboxEmailFrom = New System.Windows.Forms.TextBox()
			Me.textboxPassword = New System.Windows.Forms.TextBox()
			Me.textboxEmailTo = New System.Windows.Forms.TextBox()
			Me.buttonSend = New System.Windows.Forms.Button()
			Me.label1 = New System.Windows.Forms.Label()
			Me.label2 = New System.Windows.Forms.Label()
			Me.label3 = New System.Windows.Forms.Label()
			Me.label4 = New System.Windows.Forms.Label()
			Me.label5 = New System.Windows.Forms.Label()
			Me.textboxSubject = New System.Windows.Forms.TextBox()
			Me.openDocumentFileDialog = New System.Windows.Forms.OpenFileDialog()
			Me.panelSend = New System.Windows.Forms.Panel()
			Me.checkboxAuth = New System.Windows.Forms.CheckBox()
			Me.label6 = New System.Windows.Forms.Label()
			Me.textboxPort = New System.Windows.Forms.TextBox()
			Me.labelMessage = New System.Windows.Forms.Label()
			Me.panelSend.SuspendLayout()
			Me.SuspendLayout()
			' 
			' buttonOpenDocument
			' 
			Me.buttonOpenDocument.Location = New System.Drawing.Point(12, 12)
			Me.buttonOpenDocument.Name = "buttonOpenDocument"
			Me.buttonOpenDocument.Size = New System.Drawing.Size(75, 23)
			Me.buttonOpenDocument.TabIndex = 0
			Me.buttonOpenDocument.Text = "Open"
			Me.buttonOpenDocument.UseVisualStyleBackColor = True
'			Me.buttonOpenDocument.Click += New System.EventHandler(Me.buttonOpenDocument_Click);
			' 
			' textboxSmtp
			' 
			Me.textboxSmtp.Location = New System.Drawing.Point(124, 3)
			Me.textboxSmtp.Name = "textboxSmtp"
			Me.textboxSmtp.Size = New System.Drawing.Size(204, 20)
			Me.textboxSmtp.TabIndex = 1
			' 
			' textboxEmailFrom
			' 
			Me.textboxEmailFrom.Location = New System.Drawing.Point(124, 29)
			Me.textboxEmailFrom.Name = "textboxEmailFrom"
			Me.textboxEmailFrom.Size = New System.Drawing.Size(204, 20)
			Me.textboxEmailFrom.TabIndex = 2
			' 
			' textboxPassword
			' 
			Me.textboxPassword.Location = New System.Drawing.Point(124, 55)
			Me.textboxPassword.Name = "textboxPassword"
			Me.textboxPassword.PasswordChar = "*"c
			Me.textboxPassword.Size = New System.Drawing.Size(204, 20)
			Me.textboxPassword.TabIndex = 3
			Me.textboxPassword.UseSystemPasswordChar = True
			' 
			' textboxEmailTo
			' 
			Me.textboxEmailTo.Location = New System.Drawing.Point(124, 81)
			Me.textboxEmailTo.Name = "textboxEmailTo"
			Me.textboxEmailTo.Size = New System.Drawing.Size(204, 20)
			Me.textboxEmailTo.TabIndex = 4
			' 
			' buttonSend
			' 
			Me.buttonSend.Location = New System.Drawing.Point(6, 195)
			Me.buttonSend.Name = "buttonSend"
			Me.buttonSend.Size = New System.Drawing.Size(75, 23)
			Me.buttonSend.TabIndex = 6
			Me.buttonSend.Text = "Send"
			Me.buttonSend.UseVisualStyleBackColor = True
'			Me.buttonSend.Click += New System.EventHandler(Me.buttonSend_Click);
			' 
			' label1
			' 
			Me.label1.AutoSize = True
			Me.label1.Location = New System.Drawing.Point(4, 10)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(93, 13)
			Me.label1.TabIndex = 7
			Me.label1.Text = "smtp (smtp.mail.ru)"
			' 
			' label2
			' 
			Me.label2.AutoSize = True
			Me.label2.Location = New System.Drawing.Point(4, 36)
			Me.label2.Name = "label2"
			Me.label2.Size = New System.Drawing.Size(59, 13)
			Me.label2.TabIndex = 8
			Me.label2.Text = "Your e-mail"
			' 
			' label3
			' 
			Me.label3.AutoSize = True
			Me.label3.Location = New System.Drawing.Point(4, 62)
			Me.label3.Name = "label3"
			Me.label3.Size = New System.Drawing.Size(77, 13)
			Me.label3.TabIndex = 9
			Me.label3.Text = "Your password"
			' 
			' label4
			' 
			Me.label4.AutoSize = True
			Me.label4.Location = New System.Drawing.Point(4, 88)
			Me.label4.Name = "label4"
			Me.label4.Size = New System.Drawing.Size(82, 13)
			Me.label4.TabIndex = 10
			Me.label4.Text = "Recipient e-mail"
			' 
			' label5
			' 
			Me.label5.AutoSize = True
			Me.label5.Location = New System.Drawing.Point(4, 114)
			Me.label5.Name = "label5"
			Me.label5.Size = New System.Drawing.Size(43, 13)
			Me.label5.TabIndex = 11
			Me.label5.Text = "Subject"
			' 
			' textboxSubject
			' 
			Me.textboxSubject.Location = New System.Drawing.Point(124, 107)
			Me.textboxSubject.Name = "textboxSubject"
			Me.textboxSubject.Size = New System.Drawing.Size(204, 20)
			Me.textboxSubject.TabIndex = 5
			' 
			' openDocumentFileDialog
			' 
			Me.openDocumentFileDialog.Filter = resources.GetString("openDocumentFileDialog.Filter")
			' 
			' panelSend
			' 
			Me.panelSend.Controls.Add(Me.checkboxAuth)
			Me.panelSend.Controls.Add(Me.label6)
			Me.panelSend.Controls.Add(Me.textboxPort)
			Me.panelSend.Controls.Add(Me.labelMessage)
			Me.panelSend.Controls.Add(Me.textboxSmtp)
			Me.panelSend.Controls.Add(Me.label5)
			Me.panelSend.Controls.Add(Me.textboxEmailFrom)
			Me.panelSend.Controls.Add(Me.textboxSubject)
			Me.panelSend.Controls.Add(Me.textboxPassword)
			Me.panelSend.Controls.Add(Me.label4)
			Me.panelSend.Controls.Add(Me.textboxEmailTo)
			Me.panelSend.Controls.Add(Me.label3)
			Me.panelSend.Controls.Add(Me.buttonSend)
			Me.panelSend.Controls.Add(Me.label2)
			Me.panelSend.Controls.Add(Me.label1)
			Me.panelSend.Enabled = False
			Me.panelSend.Location = New System.Drawing.Point(12, 41)
			Me.panelSend.Name = "panelSend"
			Me.panelSend.Size = New System.Drawing.Size(340, 228)
			Me.panelSend.TabIndex = 12
			' 
			' checkboxAuth
			' 
			Me.checkboxAuth.AutoSize = True
			Me.checkboxAuth.Location = New System.Drawing.Point(124, 160)
			Me.checkboxAuth.Name = "checkboxAuth"
			Me.checkboxAuth.Size = New System.Drawing.Size(116, 17)
			Me.checkboxAuth.TabIndex = 15
			Me.checkboxAuth.Text = "Use Authentication"
			Me.checkboxAuth.UseVisualStyleBackColor = True
			' 
			' label6
			' 
			Me.label6.AutoSize = True
			Me.label6.Location = New System.Drawing.Point(4, 140)
			Me.label6.Name = "label6"
			Me.label6.Size = New System.Drawing.Size(26, 13)
			Me.label6.TabIndex = 14
			Me.label6.Text = "Port"
			' 
			' textboxPort
			' 
			Me.textboxPort.Location = New System.Drawing.Point(124, 133)
			Me.textboxPort.Name = "textboxPort"
			Me.textboxPort.Size = New System.Drawing.Size(53, 20)
			Me.textboxPort.TabIndex = 13
'			Me.textboxPort.KeyPress += New System.Windows.Forms.KeyPressEventHandler(Me.textBoxPort_KeyPress);
			' 
			' labelMessage
			' 
			Me.labelMessage.AutoSize = True
			Me.labelMessage.Font = New System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, (CByte(204)))
			Me.labelMessage.ForeColor = System.Drawing.Color.DarkGreen
			Me.labelMessage.Location = New System.Drawing.Point(121, 195)
			Me.labelMessage.Name = "labelMessage"
			Me.labelMessage.Size = New System.Drawing.Size(176, 17)
			Me.labelMessage.TabIndex = 12
			Me.labelMessage.Text = "Message sent successfully"
			Me.labelMessage.Visible = False
			' 
			' MainForm
			' 
			Me.AutoScaleDimensions = New System.Drawing.SizeF(6F, 13F)
			Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
			Me.ClientSize = New System.Drawing.Size(362, 281)
			Me.Controls.Add(Me.panelSend)
			Me.Controls.Add(Me.buttonOpenDocument)
			Me.MaximizeBox = False
			Me.Name = "MainForm"
			Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
			Me.Text = "Doc2Email"
			Me.panelSend.ResumeLayout(False)
			Me.panelSend.PerformLayout()
			Me.ResumeLayout(False)

		End Sub

		#End Region

		Private WithEvents buttonOpenDocument As System.Windows.Forms.Button
		Private textboxSmtp As System.Windows.Forms.TextBox
		Private textboxEmailFrom As System.Windows.Forms.TextBox
		Private textboxPassword As System.Windows.Forms.TextBox
		Private textboxEmailTo As System.Windows.Forms.TextBox
		Private WithEvents buttonSend As System.Windows.Forms.Button
		Private label1 As System.Windows.Forms.Label
		Private label2 As System.Windows.Forms.Label
		Private label3 As System.Windows.Forms.Label
		Private label4 As System.Windows.Forms.Label
		Private label5 As System.Windows.Forms.Label
		Private textboxSubject As System.Windows.Forms.TextBox
		Private openDocumentFileDialog As System.Windows.Forms.OpenFileDialog
		Private panelSend As System.Windows.Forms.Panel
		Private labelMessage As System.Windows.Forms.Label
		Private label6 As System.Windows.Forms.Label
		Private WithEvents textboxPort As System.Windows.Forms.TextBox
		Private checkboxAuth As System.Windows.Forms.CheckBox
	End Class
End Namespace