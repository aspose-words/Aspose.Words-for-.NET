//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Specialized;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Windows.Forms;
using Aspose.Words;
using Aspose.Words.Saving;
using System.Web.UI.WebControls;

namespace SaveHtmlAndEmailExample
{
    public partial class MainForm : Form
    {
        private string inputFileName = string.Empty;

        public MainForm()
        {
            InitializeComponent();
        }

        private void buttonOpenDocument_Click(object sender, EventArgs e)
        {
            labelMessage.Visible = false;

            try
            {
                // Prompt the user to choose the input document.
                if (openDocumentFileDialog.ShowDialog().Equals(DialogResult.OK))
                {
                    inputFileName = openDocumentFileDialog.FileName;
                    panelSend.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void buttonSend_Click(object sender, EventArgs e)
        {
            // Extract information that is needed to send email message from the user interface.
            string smtp = textboxSmtp.Text; // Your smtp server
            string emailFrom = textboxEmailFrom.Text; // Your email 
            string password = textboxPassword.Text; // Your password
            string emailTo = textboxEmailTo.Text; // Recipient email
            string subject = textboxSubject.Text; // Subject
            bool useAuth = checkboxAuth.Checked; // Use authentication 

            int port; // The port to use
            int.TryParse(textboxPort.Text, out port);

            if (port == 0)
                port = 25; // If the port was not defined it will be parsed as 0. Change to default port 25.

            try
            {
                labelMessage.Visible = false;
                buttonSend.Enabled = false;

                // Send the information required to send the e-mail.
                Send(smtp, emailFrom, password, emailTo, subject, port, useAuth, inputFileName);
                
                // Show message
                labelMessage.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            finally
            {
                // Even if an exception occurs reset the send button.
                buttonSend.Enabled = true;
            }
        }

        /// <summary>
        /// Convert document to HTML mail message and send it to recipient
        /// </summary>
        /// <param name="smtp">Smtp server</param>
        /// <param name="emailFrom">Sender e-mail</param>
        /// <param name="password">Sender password</param>
        /// <param name="emailTo">Recipient e-mail</param>
        /// <param name="subject">E-mail subject</param>
        /// <param name="port">Port to use</param>
        /// <param name="useAuth">Specify authentication</param>
        /// <param name="inputFileName">Document file name</param>
        private static void Send(string smtp, string emailFrom, string password, string emailTo, string subject, int port, bool useAuth, string inputFileName)
        {
            // Create temporary folder for Aspose.Words to store images to during export.
            string tempDir = Path.Combine(Path.GetTempPath(), "AsposeMail");
            if (!Directory.Exists(tempDir))
                Directory.CreateDirectory(tempDir);

            // Open the document.
            Document doc = new Document(inputFileName);
            
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // Specify folder where images will be saved.
            saveOptions.ImagesFolder = tempDir;
            // We want the images in the HTML to be referenced in the e-mail as attachments so add the cid prefix to the image file name.
            // This replaces what would be the path to the image with the "cid" prefix.
            saveOptions.ImagesFolderAlias = "cid:";
            // Header footers don't normally export well in HTML format so remove them.
            saveOptions.ExportHeadersFootersMode = ExportHeadersFootersMode.None;
            
            // Save the document to stream in HTML format.
            MemoryStream htmlStream = new MemoryStream();
            doc.Save(htmlStream, saveOptions);

            // Read the HTML from the stream as plain text.
            string htmlText = Encoding.UTF8.GetString(htmlStream.ToArray());
            htmlStream.Close();

            // Save the HTML into the temporary folder.
            Stream htmlFile = new FileStream(Path.Combine(tempDir, "Message.html"), FileMode.Create);
            StreamWriter htmlWriter = new StreamWriter(htmlFile);
            htmlWriter.Write(htmlText);
            htmlWriter.Close();
            htmlFile.Close();

            // Create the mail definiton and specify the appropriate settings.
            MailDefinition mail = new MailDefinition();
            mail.IsBodyHtml = true;
            mail.BodyFileName = Path.Combine(tempDir, "Message.html");
            mail.From = emailFrom;
            mail.Subject = subject;

            // Get the names of the images in the temporary folder.
            string[] fileNames = Directory.GetFiles(tempDir);

            // Add each image as an embedded object to the message.
            for (int imgIndex = 0; imgIndex < fileNames.Length; imgIndex++)
            {
                string imgFullName = fileNames[imgIndex];
                string imgName = Path.GetFileName(fileNames[imgIndex]);
                // The ID of the embedded object is the name of the image preceeded with a foward slash.
                mail.EmbeddedObjects.Add(new EmbeddedMailObject(string.Format("/{0}", imgName), imgFullName));
            }

            MailMessage message = null;

            // Create the message.
            try
            {
                message = mail.CreateMailMessage(emailTo, new ListDictionary(), new System.Web.UI.Control());
                
                // Create the SMTP client to send the message with.
                SmtpClient sender = new SmtpClient(smtp);
                
                // Set the credentials.
                sender.Credentials = new NetworkCredential(emailFrom, password);
                // Set port.
                sender.Port = port;
                // Choose to enable authentication.
                sender.EnableSsl = useAuth;
                
                // Send the e-mail message.
                sender.Send(message);
            }

            catch (Exception e)
            {
                throw e;
            }

            finally
            {
                // This frees the Message.html file if an exception occurs.
                message.Dispose();
            }

            // Delete the temp folder.
            Directory.Delete(tempDir, true);
        }

        /// <summary>
        /// This restricts the user entering anything but digits up to a certain length in the port textbox.
        /// </summary>
        private void textBoxPort_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar)) || (!char.IsControl(e.KeyChar) && ((System.Windows.Forms.TextBox)sender).Text.Length >= 5))
            {
                e.Handled = true;
            } 
        }
    }
}