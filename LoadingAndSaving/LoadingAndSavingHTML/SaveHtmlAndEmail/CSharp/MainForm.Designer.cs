// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
namespace SaveHtmlAndEmailExample
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.buttonOpenDocument = new System.Windows.Forms.Button();
            this.textboxSmtp = new System.Windows.Forms.TextBox();
            this.textboxEmailFrom = new System.Windows.Forms.TextBox();
            this.textboxPassword = new System.Windows.Forms.TextBox();
            this.textboxEmailTo = new System.Windows.Forms.TextBox();
            this.buttonSend = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.textboxSubject = new System.Windows.Forms.TextBox();
            this.openDocumentFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.panelSend = new System.Windows.Forms.Panel();
            this.checkboxAuth = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.textboxPort = new System.Windows.Forms.TextBox();
            this.labelMessage = new System.Windows.Forms.Label();
            this.panelSend.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonOpenDocument
            // 
            this.buttonOpenDocument.Location = new System.Drawing.Point(12, 12);
            this.buttonOpenDocument.Name = "buttonOpenDocument";
            this.buttonOpenDocument.Size = new System.Drawing.Size(75, 23);
            this.buttonOpenDocument.TabIndex = 0;
            this.buttonOpenDocument.Text = "Open";
            this.buttonOpenDocument.UseVisualStyleBackColor = true;
            this.buttonOpenDocument.Click += new System.EventHandler(this.buttonOpenDocument_Click);
            // 
            // textboxSmtp
            // 
            this.textboxSmtp.Location = new System.Drawing.Point(124, 3);
            this.textboxSmtp.Name = "textboxSmtp";
            this.textboxSmtp.Size = new System.Drawing.Size(204, 20);
            this.textboxSmtp.TabIndex = 1;
            // 
            // textboxEmailFrom
            // 
            this.textboxEmailFrom.Location = new System.Drawing.Point(124, 29);
            this.textboxEmailFrom.Name = "textboxEmailFrom";
            this.textboxEmailFrom.Size = new System.Drawing.Size(204, 20);
            this.textboxEmailFrom.TabIndex = 2;
            // 
            // textboxPassword
            // 
            this.textboxPassword.Location = new System.Drawing.Point(124, 55);
            this.textboxPassword.Name = "textboxPassword";
            this.textboxPassword.PasswordChar = '*';
            this.textboxPassword.Size = new System.Drawing.Size(204, 20);
            this.textboxPassword.TabIndex = 3;
            this.textboxPassword.UseSystemPasswordChar = true;
            // 
            // textboxEmailTo
            // 
            this.textboxEmailTo.Location = new System.Drawing.Point(124, 81);
            this.textboxEmailTo.Name = "textboxEmailTo";
            this.textboxEmailTo.Size = new System.Drawing.Size(204, 20);
            this.textboxEmailTo.TabIndex = 4;
            // 
            // buttonSend
            // 
            this.buttonSend.Location = new System.Drawing.Point(6, 195);
            this.buttonSend.Name = "buttonSend";
            this.buttonSend.Size = new System.Drawing.Size(75, 23);
            this.buttonSend.TabIndex = 6;
            this.buttonSend.Text = "Send";
            this.buttonSend.UseVisualStyleBackColor = true;
            this.buttonSend.Click += new System.EventHandler(this.buttonSend_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(93, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "smtp (smtp.mail.ru)";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(4, 36);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(59, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "Your e-mail";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(4, 62);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Your password";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(4, 88);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(82, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Recipient e-mail";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(4, 114);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(43, 13);
            this.label5.TabIndex = 11;
            this.label5.Text = "Subject";
            // 
            // textboxSubject
            // 
            this.textboxSubject.Location = new System.Drawing.Point(124, 107);
            this.textboxSubject.Name = "textboxSubject";
            this.textboxSubject.Size = new System.Drawing.Size(204, 20);
            this.textboxSubject.TabIndex = 5;
            // 
            // openDocumentFileDialog
            // 
            this.openDocumentFileDialog.Filter = resources.GetString("openDocumentFileDialog.Filter");
            // 
            // panelSend
            // 
            this.panelSend.Controls.Add(this.checkboxAuth);
            this.panelSend.Controls.Add(this.label6);
            this.panelSend.Controls.Add(this.textboxPort);
            this.panelSend.Controls.Add(this.labelMessage);
            this.panelSend.Controls.Add(this.textboxSmtp);
            this.panelSend.Controls.Add(this.label5);
            this.panelSend.Controls.Add(this.textboxEmailFrom);
            this.panelSend.Controls.Add(this.textboxSubject);
            this.panelSend.Controls.Add(this.textboxPassword);
            this.panelSend.Controls.Add(this.label4);
            this.panelSend.Controls.Add(this.textboxEmailTo);
            this.panelSend.Controls.Add(this.label3);
            this.panelSend.Controls.Add(this.buttonSend);
            this.panelSend.Controls.Add(this.label2);
            this.panelSend.Controls.Add(this.label1);
            this.panelSend.Enabled = false;
            this.panelSend.Location = new System.Drawing.Point(12, 41);
            this.panelSend.Name = "panelSend";
            this.panelSend.Size = new System.Drawing.Size(340, 228);
            this.panelSend.TabIndex = 12;
            // 
            // checkboxAuth
            // 
            this.checkboxAuth.AutoSize = true;
            this.checkboxAuth.Location = new System.Drawing.Point(124, 160);
            this.checkboxAuth.Name = "checkboxAuth";
            this.checkboxAuth.Size = new System.Drawing.Size(116, 17);
            this.checkboxAuth.TabIndex = 15;
            this.checkboxAuth.Text = "Use Authentication";
            this.checkboxAuth.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(4, 140);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(26, 13);
            this.label6.TabIndex = 14;
            this.label6.Text = "Port";
            // 
            // textboxPort
            // 
            this.textboxPort.Location = new System.Drawing.Point(124, 133);
            this.textboxPort.Name = "textboxPort";
            this.textboxPort.Size = new System.Drawing.Size(53, 20);
            this.textboxPort.TabIndex = 13;
            this.textboxPort.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBoxPort_KeyPress);
            // 
            // labelMessage
            // 
            this.labelMessage.AutoSize = true;
            this.labelMessage.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelMessage.ForeColor = System.Drawing.Color.DarkGreen;
            this.labelMessage.Location = new System.Drawing.Point(121, 195);
            this.labelMessage.Name = "labelMessage";
            this.labelMessage.Size = new System.Drawing.Size(176, 17);
            this.labelMessage.TabIndex = 12;
            this.labelMessage.Text = "Message sent successfully";
            this.labelMessage.Visible = false;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(362, 281);
            this.Controls.Add(this.panelSend);
            this.Controls.Add(this.buttonOpenDocument);
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Doc2Email";
            this.panelSend.ResumeLayout(false);
            this.panelSend.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonOpenDocument;
        private System.Windows.Forms.TextBox textboxSmtp;
        private System.Windows.Forms.TextBox textboxEmailFrom;
        private System.Windows.Forms.TextBox textboxPassword;
        private System.Windows.Forms.TextBox textboxEmailTo;
        private System.Windows.Forms.Button buttonSend;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textboxSubject;
        private System.Windows.Forms.OpenFileDialog openDocumentFileDialog;
        private System.Windows.Forms.Panel panelSend;
        private System.Windows.Forms.Label labelMessage;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textboxPort;
        private System.Windows.Forms.CheckBox checkboxAuth;
    }
}