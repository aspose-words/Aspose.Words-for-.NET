//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Windows;
using System.Windows.Browser;
using System.Windows.Controls;

using FirstFloor.Documents.Controls;
using FirstFloor.Documents.Extensions;
using FirstFloor.Documents.IO;

namespace FirstFloor.Documents.Aspose.Views
{
    public partial class DocumentViewer : UserControl
    {
        class UploadRequest
        {
            public string FileName { get; set; }
            public Stream FileStream { get; set; }
            public HttpWebRequest Request { get; set; }
        }

        public event EventHandler Close;
        private XpsClient client;

        public DocumentViewer()
        {
            InitializeComponent();

            this.client = new XpsClient();
            this.client.LoadXpsDocumentCompleted += client_LoadXpsDocumentCompleted;

            this.toolbar.Viewer = this.viewer;
        }

        public FixedDocumentViewer Viewer
        {
            get { return this.viewer; }
        }

        void client_LoadXpsDocumentCompleted(object sender, LoadXpsDocumentCompletedEventArgs e)
        {
            if (!e.Cancelled) {
                if (e.Error != null) {
                    ErrorWindow.ShowError(e.Error);
                }
                else {
                    this.viewer.FixedDocument = e.Document.FixedDocuments.FirstOrDefault();
                    this.toolbar.Refresh();
                }
            }
        }

        public void ClearDocument()
        {
            if (this.viewer.FixedDocument != null) {
                this.viewer.FixedDocument.Owner.Dispose();
            }
            this.viewer.FixedDocument = null;
            this.title.Text = null;
            this.toolbar.OriginalUri = null;
            this.toolbar.Refresh();
        }

        public void LoadDocument(string title, Uri uri, Uri originalUri)
        {
            ClearDocument();

            this.title.Text = title;
            this.toolbar.OriginalUri = originalUri;
            this.loading.Visibility = Visibility.Visible;
            this.status.Text = "Loading...Please Wait";

            var webClient = new WebClient();
            webClient.OpenReadCompleted += (o, e) => {
                if (!e.Cancelled){
                    if (e.Error != null) {
                        ErrorWindow.ShowError(e.Error);
                    }
                    else {
                        LoadDocument(e.Result);
                    }
                }
                this.loading.Visibility = Visibility.Collapsed;
            };

            webClient.OpenReadAsync(uri);
        }

        public void LoadLocalDocument(FileInfo file)
        {
            try {
                var uri = new Uri(HtmlPage.Document.DocumentUri, "ConvertToXps.ashx");
                var request = (HttpWebRequest)WebRequest.Create(uri);
                var uploadRequest = new UploadRequest() {
                    FileName = file.Name,
                    FileStream = file.OpenRead(),
                    Request = request
                };
                request.Method = "POST";
                request.BeginGetRequestStream(OnGetRequestStream, uploadRequest);

                ClearDocument();
                this.loading.Visibility = Visibility.Visible;
                this.status.Text = "Sending document...Please Wait";
                this.title.Text = uploadRequest.FileName;
            }
            catch (Exception ex) {
                ErrorWindow.ShowError(ex);
                this.loading.Visibility = Visibility.Collapsed;
            }
        }

        private void LoadDocument(Stream stream)
        {
            var reader = new SharpZipPackageReader(stream);

            var settings = new LoadXpsDocumentSettings() {
                IncludeProperties = false,
                IncludeDocumentStructures = false,
                IncludeAnnotations = false
            };

            this.client.LoadXpsDocumentAsync(reader, settings);
        }

        private void OnGetRequestStream(IAsyncResult result)
        {
            var uploadRequest = (UploadRequest)result.AsyncState;
            try {
                using (var stream = uploadRequest.Request.EndGetRequestStream(result)) {
                    var buffer = new byte[4096];
                    int bytesRead;

                    while ((bytesRead = uploadRequest.FileStream.Read(buffer, 0, buffer.Length)) != 0) {
                        stream.Write(buffer, 0, bytesRead);
                    }
                }

                Dispatcher.BeginInvoke(() => {
                    this.status.Text = "Receiving XPS...Please Wait";
                });
                
                uploadRequest.Request.BeginGetResponse(OnGetResponse, uploadRequest);
            }
            catch (Exception e) {
                Dispatcher.BeginInvoke(() => {
                    this.loading.Visibility = Visibility.Collapsed;
                    ErrorWindow.ShowError(e);
                });
            }
            finally {
                uploadRequest.FileStream.Close();
            }
        }

        private void OnGetResponse(IAsyncResult result)
        {
            try {
                var uploadRequest = (UploadRequest)result.AsyncState;
                var response = (HttpWebResponse)uploadRequest.Request.EndGetResponse(result);

                Dispatcher.BeginInvoke(() => {
                    LoadDocument(response.GetResponseStream());
                    this.loading.Visibility = Visibility.Collapsed;
                });

            }
            catch (Exception e) {
                Dispatcher.BeginInvoke(() => {
                    this.loading.Visibility = Visibility.Collapsed;
                    ErrorWindow.ShowError(e);
                });
            }
        }

        private void close_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Host.Content.IsFullScreen = false;
            if (Close != null) {
                Close(this, EventArgs.Empty);
            }
        }

        private void viewer_LinkClick(object sender, LinkClickEventArgs e)
        {
            if (e.NavigateUri.IsAbsoluteUri && (e.NavigateUri.Scheme == Uri.UriSchemeHttp || e.NavigateUri.Scheme == Uri.UriSchemeHttps)) {
                HtmlPage.Window.Navigate(e.NavigateUri, "_blank");
            }
        }
    }
}
