using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.IO;
using System.Windows.Forms;
using Aspose.Words;

namespace Aspose.Words_Metadata_Cleaner
{
    public partial class ThisAddIn
    {
        public bool EnableAsposeWordsMetadata;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Enable the Add-in on start of Outlook.
            EnableAsposeWordsMetadata = true;
            try
            {
                // Apply the Aspose license if it exists.
                string LicenceFilePath = String.IsNullOrEmpty(Properties.Settings.Default.AsposeLicense) ? "" : Properties.Settings.Default.AsposeLicense;
                try
                {
                    if (LicenceFilePath != "")
                    {
                        if (File.Exists(LicenceFilePath))
                        {
                            Words.License Lic = new Words.License();
                            Lic.SetLicense(LicenceFilePath);
                        }
                        else
                        {
                            MessageBox.Show("Aspose License file doesnot exist on the location");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Not able to set license. Application will run in demo mode. Error: " + ex.Message);
                }

                // Start a plugin on email sent.
                Application.ItemSend += new
            Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Error");
            }
        }

        // This method is called every time an email is sent.
        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            if (EnableAsposeWordsMetadata)
            {
                try
                {
                    Outlook.MailItem thisEmail = Item as Outlook.MailItem;
                    bool flgPassword = false;
                    for (int i = thisEmail.Attachments.Count; i > 0; i--)
                    {
                        Outlook.Attachment attachment = thisEmail.Attachments[i];

                        // save attachment in temp location
                        int attachmentIndex = attachment.Index;
                        string tempPath = Path.GetTempPath();
                        string tempFileName = tempPath + attachment.FileName;
                        attachment.SaveAsFile(tempFileName);

                        // Check the file format for word documents
                        FileFormatInfo info = FileFormatUtil.DetectFileFormat(tempFileName);
                        bool wordAttachment = false;

                        switch (info.LoadFormat)
                        {
                            case LoadFormat.Doc:
                            case LoadFormat.Dot:
                            case LoadFormat.Docx:
                            case LoadFormat.Docm:
                            case LoadFormat.Dotx:
                            case LoadFormat.Dotm:
                            case LoadFormat.FlatOpc:
                            case LoadFormat.Rtf:
                            case LoadFormat.WordML:
                            case LoadFormat.Html:
                            case LoadFormat.Mhtml:
                            case LoadFormat.Odt:
                            case LoadFormat.Ott:
                            case LoadFormat.DocPreWord60:
                                wordAttachment = true;
                                break;
                        }

                        // If word Attachment is found
                        if (wordAttachment)
                        {
                            try
                            {
                                Aspose.Words.Document doc = new Words.Document(tempFileName);

                                // Remove if there is any protection on the document.
                                ProtectionType protection = doc.ProtectionType;
                                if (protection != ProtectionType.NoProtection)
                                    doc.Unprotect();

                                // Remove all built-in and Custom Properties.
                                doc.CustomDocumentProperties.Clear();
                                doc.BuiltInDocumentProperties.Clear();

                                // Password will be removed if the document is password protected.
                                if (protection != ProtectionType.NoProtection)
                                    doc.Protect(protection);

                                // Save the file back to temp location.
                                doc.Save(tempFileName);

                                // Replace the original attachment.
                                thisEmail.Attachments.Remove(attachmentIndex);
                                thisEmail.Attachments.Add(tempFileName, missing, attachmentIndex, missing);
                            }
                            catch (Words.IncorrectPasswordException ex)
                            {
                                flgPassword = true;
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                        }
                        // Delete file from temp folder.
                        if (File.Exists(tempFileName))
                            File.Delete(tempFileName);
                    }
                    if (flgPassword)
                    {
                        MessageBox.Show("Password protected documents cannot be processed for Metadata cleaning");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message, "Error");
                    Cancel = true;
                }
            }
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
