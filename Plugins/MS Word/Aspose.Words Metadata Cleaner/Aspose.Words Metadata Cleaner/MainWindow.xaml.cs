using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using AsposeWords = Aspose.Words;

namespace Aspose.Words_Metadata_Cleaner
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.AddExtension = true;
            openFileDialog.Filter = "Microsoft Word Documents (*.doc, *.docx)|*.doc;*.docx";
            openFileDialog.Multiselect = true;
            if (openFileDialog.ShowDialog() == true)
            {
                TXT_FileNames.Text = "";
                string[] FileNames = openFileDialog.FileNames;
                foreach (string FileName in FileNames)
                    TXT_FileNames.Text += FileName + Environment.NewLine;
                BTN_Clean.IsEnabled = true;
            }
            else
                BTN_Clean.IsEnabled = false;
        }

        private void Clean_Click(object sender, RoutedEventArgs e)
        {
            string Files = TXT_FileNames.Text;
            string[] FileNames = Files.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            List<string> ProcessedFiles = new List<string>();
            int Total = 0;
            int Successful = 0;
            Total = FileNames.Length;
            foreach (string ThisFile in FileNames)
            {
                bool cleaned = false;
                string Message = "";
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
                        cleaned = true;
                        Successful++;
                    }
                    catch (Words.IncorrectPasswordException)
                    {
                        Message = "Password Protected File";
                    }
                    catch (Exception ex)
                    {
                        Message = "Error: " + ex.Message;
                    }
                }
                else
                    Message = "Not a Word Document";
                if (cleaned)
                    ProcessedFiles.Add(ThisFile + " (Metadata Cleaned Successfully)");
                else
                    ProcessedFiles.Add(ThisFile + " (" + Message + ")");
            }
            TXT_FileNames.Text = "";
            foreach (string FileName in ProcessedFiles)
                TXT_FileNames.Text += FileName + Environment.NewLine;
            BTN_Clean.IsEnabled = false;
            LBL_Status.Content = "Status:" + Environment.NewLine;
            LBL_Status.Content += "Total Files: " + Total.ToString() + Environment.NewLine;
            LBL_Status.Content += "Cleaned: " + Successful.ToString();
        }

        private void ApplyLicense_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.AddExtension = true;
            openFileDialog.Filter = "Aspose License File (*.lic)|*.lic";
            if (openFileDialog.ShowDialog() == true)
            {
                string LicenseFile = openFileDialog.FileName;
                try
                {
                    if (LicenseFile != "")
                    {
                        if (File.Exists(LicenseFile))
                        {
                          AsposeWords.  License Lic = new AsposeWords.License();
                            Lic.SetLicense(LicenseFile);
                            if (Lic.IsLicensed)
                            {
                                LBL_License.Content = "Licensed";
                                BTN_ApplyLicense.IsEnabled = true;
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    LBL_License.Content = "License not applied";
                }
            }
        }
    }
}
