namespace Aspose.Words_MetadataCleaner
{
    partial class MetadataCleaner
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MetadataCleaner));
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.PNL_License = new System.Windows.Forms.Panel();
            this.BTN_ApplyLicense = new System.Windows.Forms.Button();
            this.LBL_AsposeLink = new System.Windows.Forms.LinkLabel();
            this.LBL_LicenseDetails = new System.Windows.Forms.Label();
            this.LBL_ApplyLicense = new System.Windows.Forms.Label();
            this.PNL_Select = new System.Windows.Forms.Panel();
            this.BTN_BrowseFiles = new System.Windows.Forms.Button();
            this.LBL_SelectDescription = new System.Windows.Forms.Label();
            this.LBL_Select = new System.Windows.Forms.Label();
            this.PNL_Clean = new System.Windows.Forms.Panel();
            this.BTN_Clean = new System.Windows.Forms.Button();
            this.LBL_CleanDescription = new System.Windows.Forms.Label();
            this.LBL_CleanMetadata = new System.Windows.Forms.Label();
            this.PNL_Status = new System.Windows.Forms.Panel();
            this.LBL_Cleaned = new System.Windows.Forms.Label();
            this.LBL_CleanedFiles = new System.Windows.Forms.Label();
            this.LBL_Total = new System.Windows.Forms.Label();
            this.LBL_TotalDocument = new System.Windows.Forms.Label();
            this.LBL_LicenseStatus = new System.Windows.Forms.Label();
            this.LBL_LicenseStatusLabel = new System.Windows.Forms.Label();
            this.LBL_Statics = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.LBL_Error = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.PNL_License.SuspendLayout();
            this.PNL_Select.SuspendLayout();
            this.PNL_Clean.SuspendLayout();
            this.PNL_Status.SuspendLayout();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.pictureBox1.Image = global::Aspose.Words_MetadataCleaner.Properties.Resources.asposeLogo;
            this.pictureBox1.Location = new System.Drawing.Point(0, 0);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(783, 50);
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // PNL_License
            // 
            this.PNL_License.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.PNL_License.Controls.Add(this.BTN_ApplyLicense);
            this.PNL_License.Controls.Add(this.LBL_AsposeLink);
            this.PNL_License.Controls.Add(this.LBL_LicenseDetails);
            this.PNL_License.Controls.Add(this.LBL_ApplyLicense);
            this.PNL_License.Location = new System.Drawing.Point(10, 55);
            this.PNL_License.Name = "PNL_License";
            this.PNL_License.Size = new System.Drawing.Size(760, 100);
            this.PNL_License.TabIndex = 1;
            // 
            // BTN_ApplyLicense
            // 
            this.BTN_ApplyLicense.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BTN_ApplyLicense.Location = new System.Drawing.Point(10, 50);
            this.BTN_ApplyLicense.Name = "BTN_ApplyLicense";
            this.BTN_ApplyLicense.Size = new System.Drawing.Size(120, 30);
            this.BTN_ApplyLicense.TabIndex = 4;
            this.BTN_ApplyLicense.Text = "Apply License";
            this.BTN_ApplyLicense.UseVisualStyleBackColor = true;
            this.BTN_ApplyLicense.Click += new System.EventHandler(this.BTN_ApplyLicense_Click);
            // 
            // LBL_AsposeLink
            // 
            this.LBL_AsposeLink.AutoSize = true;
            this.LBL_AsposeLink.LinkBehavior = System.Windows.Forms.LinkBehavior.AlwaysUnderline;
            this.LBL_AsposeLink.Location = new System.Drawing.Point(555, 30);
            this.LBL_AsposeLink.Name = "LBL_AsposeLink";
            this.LBL_AsposeLink.Size = new System.Drawing.Size(171, 13);
            this.LBL_AsposeLink.TabIndex = 3;
            this.LBL_AsposeLink.TabStop = true;
            this.LBL_AsposeLink.Text = "http://www.aspose.com/purchase";
            this.LBL_AsposeLink.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LBL_AsposeLink_LinkClicked);
            // 
            // LBL_LicenseDetails
            // 
            this.LBL_LicenseDetails.AutoSize = true;
            this.LBL_LicenseDetails.Location = new System.Drawing.Point(30, 30);
            this.LBL_LicenseDetails.Name = "LBL_LicenseDetails";
            this.LBL_LicenseDetails.Size = new System.Drawing.Size(530, 13);
            this.LBL_LicenseDetails.TabIndex = 2;
            this.LBL_LicenseDetails.Text = "Once you are happy with your evaluation of Aspose.Words, you can purchase a licen" +
    "se at the Aspose website:";
            // 
            // LBL_ApplyLicense
            // 
            this.LBL_ApplyLicense.AutoSize = true;
            this.LBL_ApplyLicense.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LBL_ApplyLicense.Location = new System.Drawing.Point(10, 5);
            this.LBL_ApplyLicense.Name = "LBL_ApplyLicense";
            this.LBL_ApplyLicense.Size = new System.Drawing.Size(136, 20);
            this.LBL_ApplyLicense.TabIndex = 1;
            this.LBL_ApplyLicense.Text = "Aspose License";
            // 
            // PNL_Select
            // 
            this.PNL_Select.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.PNL_Select.Controls.Add(this.BTN_BrowseFiles);
            this.PNL_Select.Controls.Add(this.LBL_SelectDescription);
            this.PNL_Select.Controls.Add(this.LBL_Select);
            this.PNL_Select.Location = new System.Drawing.Point(10, 175);
            this.PNL_Select.Name = "PNL_Select";
            this.PNL_Select.Size = new System.Drawing.Size(760, 100);
            this.PNL_Select.TabIndex = 2;
            // 
            // BTN_BrowseFiles
            // 
            this.BTN_BrowseFiles.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BTN_BrowseFiles.Location = new System.Drawing.Point(10, 50);
            this.BTN_BrowseFiles.Name = "BTN_BrowseFiles";
            this.BTN_BrowseFiles.Size = new System.Drawing.Size(120, 30);
            this.BTN_BrowseFiles.TabIndex = 4;
            this.BTN_BrowseFiles.Text = "Browse Files";
            this.BTN_BrowseFiles.UseVisualStyleBackColor = true;
            this.BTN_BrowseFiles.Click += new System.EventHandler(this.BTN_BrowseFiles_Click);
            // 
            // LBL_SelectDescription
            // 
            this.LBL_SelectDescription.AutoSize = true;
            this.LBL_SelectDescription.Location = new System.Drawing.Point(30, 30);
            this.LBL_SelectDescription.Name = "LBL_SelectDescription";
            this.LBL_SelectDescription.Size = new System.Drawing.Size(463, 13);
            this.LBL_SelectDescription.TabIndex = 2;
            this.LBL_SelectDescription.Text = "Select single or multiple documents to clean Metadata. It support All Word proces" +
    "sing documents";
            // 
            // LBL_Select
            // 
            this.LBL_Select.AutoSize = true;
            this.LBL_Select.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LBL_Select.Location = new System.Drawing.Point(10, 5);
            this.LBL_Select.Name = "LBL_Select";
            this.LBL_Select.Size = new System.Drawing.Size(103, 20);
            this.LBL_Select.TabIndex = 1;
            this.LBL_Select.Text = "Select Files";
            // 
            // PNL_Clean
            // 
            this.PNL_Clean.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.PNL_Clean.Controls.Add(this.BTN_Clean);
            this.PNL_Clean.Controls.Add(this.LBL_CleanDescription);
            this.PNL_Clean.Controls.Add(this.LBL_CleanMetadata);
            this.PNL_Clean.Location = new System.Drawing.Point(10, 290);
            this.PNL_Clean.Name = "PNL_Clean";
            this.PNL_Clean.Size = new System.Drawing.Size(760, 100);
            this.PNL_Clean.TabIndex = 5;
            // 
            // BTN_Clean
            // 
            this.BTN_Clean.Enabled = false;
            this.BTN_Clean.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BTN_Clean.Location = new System.Drawing.Point(10, 50);
            this.BTN_Clean.Name = "BTN_Clean";
            this.BTN_Clean.Size = new System.Drawing.Size(120, 30);
            this.BTN_Clean.TabIndex = 4;
            this.BTN_Clean.Text = "Clean Files";
            this.BTN_Clean.UseVisualStyleBackColor = true;
            this.BTN_Clean.Click += new System.EventHandler(this.BTN_Clean_Click);
            // 
            // LBL_CleanDescription
            // 
            this.LBL_CleanDescription.AutoSize = true;
            this.LBL_CleanDescription.Location = new System.Drawing.Point(30, 30);
            this.LBL_CleanDescription.Name = "LBL_CleanDescription";
            this.LBL_CleanDescription.Size = new System.Drawing.Size(297, 13);
            this.LBL_CleanDescription.TabIndex = 2;
            this.LBL_CleanDescription.Text = "Clean all Built-in and Custom properties from the selected files.";
            // 
            // LBL_CleanMetadata
            // 
            this.LBL_CleanMetadata.AutoSize = true;
            this.LBL_CleanMetadata.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LBL_CleanMetadata.Location = new System.Drawing.Point(10, 5);
            this.LBL_CleanMetadata.Name = "LBL_CleanMetadata";
            this.LBL_CleanMetadata.Size = new System.Drawing.Size(136, 20);
            this.LBL_CleanMetadata.TabIndex = 1;
            this.LBL_CleanMetadata.Text = "Clean Metadata";
            // 
            // PNL_Status
            // 
            this.PNL_Status.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.PNL_Status.Controls.Add(this.label1);
            this.PNL_Status.Controls.Add(this.LBL_Error);
            this.PNL_Status.Controls.Add(this.LBL_Cleaned);
            this.PNL_Status.Controls.Add(this.LBL_CleanedFiles);
            this.PNL_Status.Controls.Add(this.LBL_Total);
            this.PNL_Status.Controls.Add(this.LBL_TotalDocument);
            this.PNL_Status.Controls.Add(this.LBL_LicenseStatus);
            this.PNL_Status.Controls.Add(this.LBL_LicenseStatusLabel);
            this.PNL_Status.Controls.Add(this.LBL_Statics);
            this.PNL_Status.Location = new System.Drawing.Point(10, 410);
            this.PNL_Status.Name = "PNL_Status";
            this.PNL_Status.Size = new System.Drawing.Size(760, 120);
            this.PNL_Status.TabIndex = 6;
            // 
            // LBL_Cleaned
            // 
            this.LBL_Cleaned.AutoSize = true;
            this.LBL_Cleaned.Location = new System.Drawing.Point(116, 70);
            this.LBL_Cleaned.Name = "LBL_Cleaned";
            this.LBL_Cleaned.Size = new System.Drawing.Size(10, 13);
            this.LBL_Cleaned.TabIndex = 7;
            this.LBL_Cleaned.Text = "-";
            // 
            // LBL_CleanedFiles
            // 
            this.LBL_CleanedFiles.AutoSize = true;
            this.LBL_CleanedFiles.Location = new System.Drawing.Point(30, 70);
            this.LBL_CleanedFiles.Name = "LBL_CleanedFiles";
            this.LBL_CleanedFiles.Size = new System.Drawing.Size(73, 13);
            this.LBL_CleanedFiles.TabIndex = 6;
            this.LBL_CleanedFiles.Text = "Cleaned Files:";
            // 
            // LBL_Total
            // 
            this.LBL_Total.AutoSize = true;
            this.LBL_Total.Location = new System.Drawing.Point(116, 50);
            this.LBL_Total.Name = "LBL_Total";
            this.LBL_Total.Size = new System.Drawing.Size(10, 13);
            this.LBL_Total.TabIndex = 5;
            this.LBL_Total.Text = "-";
            // 
            // LBL_TotalDocument
            // 
            this.LBL_TotalDocument.AutoSize = true;
            this.LBL_TotalDocument.Location = new System.Drawing.Point(30, 50);
            this.LBL_TotalDocument.Name = "LBL_TotalDocument";
            this.LBL_TotalDocument.Size = new System.Drawing.Size(58, 13);
            this.LBL_TotalDocument.TabIndex = 4;
            this.LBL_TotalDocument.Text = "Total Files:";
            // 
            // LBL_LicenseStatus
            // 
            this.LBL_LicenseStatus.AutoSize = true;
            this.LBL_LicenseStatus.Location = new System.Drawing.Point(116, 30);
            this.LBL_LicenseStatus.Name = "LBL_LicenseStatus";
            this.LBL_LicenseStatus.Size = new System.Drawing.Size(99, 13);
            this.LBL_LicenseStatus.TabIndex = 3;
            this.LBL_LicenseStatus.Text = "License not applied";
            // 
            // LBL_LicenseStatusLabel
            // 
            this.LBL_LicenseStatusLabel.AutoSize = true;
            this.LBL_LicenseStatusLabel.Location = new System.Drawing.Point(30, 30);
            this.LBL_LicenseStatusLabel.Name = "LBL_LicenseStatusLabel";
            this.LBL_LicenseStatusLabel.Size = new System.Drawing.Size(80, 13);
            this.LBL_LicenseStatusLabel.TabIndex = 2;
            this.LBL_LicenseStatusLabel.Text = "License Status:";
            // 
            // LBL_Statics
            // 
            this.LBL_Statics.AutoSize = true;
            this.LBL_Statics.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LBL_Statics.Location = new System.Drawing.Point(10, 5);
            this.LBL_Statics.Name = "LBL_Statics";
            this.LBL_Statics.Size = new System.Drawing.Size(62, 20);
            this.LBL_Statics.TabIndex = 1;
            this.LBL_Statics.Text = "Status";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(116, 92);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(10, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "-";
            // 
            // LBL_Error
            // 
            this.LBL_Error.AutoSize = true;
            this.LBL_Error.Location = new System.Drawing.Point(30, 92);
            this.LBL_Error.Name = "LBL_Error";
            this.LBL_Error.Size = new System.Drawing.Size(37, 13);
            this.LBL_Error.TabIndex = 8;
            this.LBL_Error.Text = "Errors:";
            // 
            // MetadataCleaner
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 561);
            this.Controls.Add(this.PNL_Status);
            this.Controls.Add(this.PNL_Clean);
            this.Controls.Add(this.PNL_Select);
            this.Controls.Add(this.PNL_License);
            this.Controls.Add(this.pictureBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MetadataCleaner";
            this.Text = "Aspose.Words Metadata Cleaner";
            this.Load += new System.EventHandler(this.MetadataCleaner_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.PNL_License.ResumeLayout(false);
            this.PNL_License.PerformLayout();
            this.PNL_Select.ResumeLayout(false);
            this.PNL_Select.PerformLayout();
            this.PNL_Clean.ResumeLayout(false);
            this.PNL_Clean.PerformLayout();
            this.PNL_Status.ResumeLayout(false);
            this.PNL_Status.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Panel PNL_License;
        private System.Windows.Forms.Label LBL_ApplyLicense;
        private System.Windows.Forms.Label LBL_LicenseDetails;
        private System.Windows.Forms.LinkLabel LBL_AsposeLink;
        private System.Windows.Forms.Button BTN_ApplyLicense;
        private System.Windows.Forms.Panel PNL_Select;
        private System.Windows.Forms.Button BTN_BrowseFiles;
        private System.Windows.Forms.Label LBL_SelectDescription;
        private System.Windows.Forms.Label LBL_Select;
        private System.Windows.Forms.Panel PNL_Clean;
        private System.Windows.Forms.Button BTN_Clean;
        private System.Windows.Forms.Label LBL_CleanDescription;
        private System.Windows.Forms.Label LBL_CleanMetadata;
        private System.Windows.Forms.Panel PNL_Status;
        private System.Windows.Forms.Label LBL_LicenseStatus;
        private System.Windows.Forms.Label LBL_LicenseStatusLabel;
        private System.Windows.Forms.Label LBL_Statics;
        private System.Windows.Forms.Label LBL_Cleaned;
        private System.Windows.Forms.Label LBL_CleanedFiles;
        private System.Windows.Forms.Label LBL_Total;
        private System.Windows.Forms.Label LBL_TotalDocument;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label LBL_Error;
    }
}

