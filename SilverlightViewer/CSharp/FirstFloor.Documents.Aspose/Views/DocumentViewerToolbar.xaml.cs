//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Browser;
using System.Windows.Controls;
using System.Windows.Input;

using FirstFloor.Documents.Controls;

namespace FirstFloor.Documents.Aspose.Views
{
    public partial class DocumentViewerToolbar : UserControl
    {
        private FixedDocumentViewer viewer;
        private Uri originalUri;

        public DocumentViewerToolbar()
        {
            InitializeComponent();

            var modes = ViewMode.GetDefaultItems();
            this.viewMode.ItemsSource = modes;
            this.viewMode.SelectedItem = modes.First(m => m.Scale == 1);
            this.IsEnabled = false;
        }

        public FixedDocumentViewer Viewer
        {
            get { return this.viewer; }
            set
            {
                this.viewer = value;
                this.viewer.ViewMode = (ViewMode)this.viewMode.SelectedItem;
                this.IsEnabled = this.viewer != null;

                if (this.viewer != null) {
                    this.viewer.PageNumberChanged += viewer_PageNumberChanged;
                }
            }
        }

        private void viewer_PageNumberChanged(object sender, PageNumberChangedEventArgs e)
        {
            SelectPage(e.PageNumber);
        }

        public Uri OriginalUri
        {
            get { return this.originalUri; }
            set
            {
                this.originalUri = value;
                this.download.Visibility = value != null ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        public void Refresh()
        {
            if (this.viewer != null) {
                SelectPage(1);
                this.IsEnabled = this.viewer.PageCount > 0;
            }
        }

        private void SelectPage(int pageNumber)
        {
            var pageCount = this.viewer.PageCount;
            pageNumber = Math.Max(pageCount > 0 ? 1 : 0, Math.Min(pageNumber, pageCount));

            this.viewer.PageNumber = pageNumber;
            this.page.Text = string.Format("{0}", pageNumber);
            this.pages.Text = string.Format("/ {0}", pageCount);
            this.next.IsEnabled = pageNumber < pageCount;
            this.prev.IsEnabled = pageNumber > 1;
        }

        private void prev_Click(object sender, RoutedEventArgs e)
        {
            SelectPage(this.viewer.PageNumber - 1);
        }

        private void next_Click(object sender, RoutedEventArgs e)
        {
            SelectPage(this.viewer.PageNumber + 1);
        }

        private void page_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter) {
                try {
                    int pageNumber = int.Parse(page.Text, CultureInfo.InvariantCulture);
                    SelectPage(pageNumber);
                }
                catch (FormatException) {
                    SelectPage(this.viewer.PageNumber);
                }
            }
        }

        private void viewMode_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (this.viewer != null) {
                this.viewer.ViewMode = (ViewMode)viewMode.SelectedItem;

                // set focus back to viewer
                this.viewer.Focus();
            }
        }

        private void screen_Click(object sender, RoutedEventArgs e)
        {
            App.Current.Host.Content.IsFullScreen = !App.Current.Host.Content.IsFullScreen;
        }

        private void download_Click(object sender, RoutedEventArgs e)
        {
            HtmlPage.Window.Navigate(this.originalUri, "_blank");
        }
    }
}
