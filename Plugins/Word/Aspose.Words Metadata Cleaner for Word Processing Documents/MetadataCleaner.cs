using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using AsposeWords = Aspose.Words;

namespace Aspose.Words_MetadataCleaner
{
    public partial class MetadataCleaner : Form
    {
        List<string> InputFiles = new List<string>(); 
        public MetadataCleaner()
        {
            InitializeComponent();
        }

        private void MetadataCleaner_Load(object sender, EventArgs e)
        {
            LBL_AsposeLink.Links.Add(new LinkLabel.Link(0, 30, "http://www.aspose.com/purchase"));
        }

        private void BTN_BrowseFiles_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.AddExtension = true;
            openFileDialog.Filter = "Word Processing Documents|*.doc;*.docx;*.dot;*.docm;*.dotx;*.dotm;*.rtf;*.odt;*.ott|All Files (*.*)|*.*";
            openFileDialog.Multiselect = true;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                InputFiles.Clear();
                string[] FileNames = openFileDialog.FileNames;
                foreach (string FileName in FileNames)
                    InputFiles.Add(FileName);
                BTN_Clean.Enabled = true;
                LBL_Total.Text = InputFiles.Count.ToString();
            }
            else
                BTN_Clean.Enabled = false;
        }

        private void BTN_ApplyLicense_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.AddExtension = true;
            openFileDialog.Filter = "Aspose License File (*.lic)|*.lic";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string LicenseFile = openFileDialog.FileName;
                try
                {
                    if (LicenseFile != "")
                    {
                        if (File.Exists(LicenseFile))
                        {
                            AsposeWords.License Lic = new AsposeWords.License();
                            Lic.SetLicense(LicenseFile);
                            if (Lic.IsLicensed)
                            {
                                LBL_LicenseStatus.Text = "Licensed";
                                BTN_ApplyLicense.Enabled = false;
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    LBL_LicenseStatus.Text = "License not applied";
                }
            }
        }

        private void BTN_Clean_Click(object sender, EventArgs e)
        {
            int cleanedCount = 0;
            bool passwordFlg = false;
            bool errorFlg = false;
            string Message = "";
            foreach (string ThisFile in InputFiles)
            {
                AsposeWords.FileFormatInfo info = AsposeWords.FileFormatUtil.DetectFileFormat(ThisFile);
                bool WordAttachment = false;

                switch (info.LoadFormat)
                {
                    case AsposeWords.LoadFormat.Doc:
                    case AsposeWords.LoadFormat.Dot:
                    case AsposeWords.LoadFormat.Docx:
                    case AsposeWords.LoadFormat.Docm:
                    case AsposeWords.LoadFormat.Dotx:
                    case AsposeWords.LoadFormat.Dotm:
                    case AsposeWords.LoadFormat.FlatOpc:
                    case AsposeWords.LoadFormat.Rtf:
                    case AsposeWords.LoadFormat.WordML:
                    case AsposeWords.LoadFormat.Html:
                    case AsposeWords.LoadFormat.Mhtml:
                    case AsposeWords.LoadFormat.Odt:
                    case AsposeWords.LoadFormat.Ott:
                    case AsposeWords.LoadFormat.DocPreWord97:
                        WordAttachment = true;
                        break;
                    default:
                        WordAttachment = false;
                        break;
                }

                // If word Attachment is found
                if (WordAttachment)
                {
                    try
                    {
                        AsposeWords.Document doc = new AsposeWords.Document(ThisFile);

                        // Remove if there is any protection on the document
                        AsposeWords.ProtectionType protection = doc.ProtectionType;
                        if (protection != AsposeWords.ProtectionType.NoProtection)
                            doc.Unprotect();

                        // Remove all built-in and Custom Properties
                        doc.CustomDocumentProperties.Clear();
                        doc.BuiltInDocumentProperties.Clear();

                        // Password will be removed if the document is password protected.
                        if (protection != AsposeWords.ProtectionType.NoProtection)
                            doc.Protect(protection);

                        // Save the file back to temp location
                        doc.Save(ThisFile);
                        cleanedCount++;
                    }
                    catch (Words.IncorrectPasswordException)
                    {
                        passwordFlg = true;
                        Message = "Password protected files cannot be cleaned";
                    }
                    catch (Exception ex)
                    {
                        errorFlg = true;
                        Message = "Error: " + ex.Message;
                    }
                }
                else
                    Message = "Not a Word Document";
            }
            if (passwordFlg)
                LBL_Error.Text = Message;
            if (errorFlg)
                LBL_Error.Text = Message;
            BTN_Clean.Enabled = false;
            LBL_Cleaned.Text = cleanedCount.ToString();
        }

        private void LBL_AsposeLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string target = e.Link.LinkData as string;
            System.Diagnostics.Process.Start(target);
        }

        
    }
}
