//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.Windows.Navigation;
using System.IO;
using System.Windows.Browser;

namespace FirstFloor.Documents.Aspose.Views
{
    public partial class Demo : Page
    {
        public Demo()
        {
            InitializeComponent();

            this.explorer.DocumentSelected += new EventHandler<DocumentSelectedEventArgs>(explorer_DocumentSelected);
        }

        private void explorer_DocumentSelected(object sender, DocumentSelectedEventArgs e)
        {
            this.PageScrollViewer.Visibility = Visibility.Collapsed;
            this.viewer.Visibility = Visibility.Visible;

            this.viewer.LoadDocument(e.Document.Name, e.Document.XpsLocation, e.Document.OriginalLocation);
        }

        // Executes when the user navigates to this page.
        protected override void OnNavigatedTo(NavigationEventArgs e)
        {
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog();
            dlg.Multiselect = false;
            dlg.Filter = "Word documents (*.doc,*.docx)|*.doc;*.docx|WordML documents (*.xml)|*.xml|OpenDocument Text (*.odt)|*.odt|HTML pages (*.htm,*.html)|*.htm;*.html|RTF documents (*.rtf)|*.rtf|All files (*.*)|*.*";

            if (true == dlg.ShowDialog()) {
                if (dlg.File.Length > 2 << 18) {
                    MessageBox.Show("The selected document is too large. This demo limits the file size to 512KB.\n\nPlease select a smaller document.", "Document too large", MessageBoxButton.OK);
                    return;
                }
                this.PageScrollViewer.Visibility = Visibility.Collapsed;
                this.viewer.Visibility = Visibility.Visible;
                this.viewer.ClearDocument();

                this.viewer.LoadLocalDocument(dlg.File);
            }
        }

        private void viewer_Close(object sender, EventArgs e)
        {
            this.PageScrollViewer.Visibility = Visibility.Visible;
            this.viewer.Visibility = Visibility.Collapsed;
        }
    }
}
