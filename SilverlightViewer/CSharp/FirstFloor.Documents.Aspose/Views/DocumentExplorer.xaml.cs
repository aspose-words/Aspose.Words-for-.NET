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
using System.Windows.Browser;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using System.Xml.Linq;

namespace FirstFloor.Documents.Aspose.Views
{
    public partial class DocumentExplorer : UserControl
    {
        public event EventHandler<DocumentSelectedEventArgs> DocumentSelected;

        public DocumentExplorer()
        {
            InitializeComponent();

            var documentsUri = new Uri(HtmlPage.Document.DocumentUri, "Documents/Documents.xml");

            var client = new WebClient();
            client.OpenReadCompleted += (o, e) => {
                if (!e.Cancelled) {
                    if (e.Error != null) {
                        ErrorWindow.ShowError(e.Error);
                    }
                    else {
                        using (e.Result) {
                            var doc = XDocument.Load(e.Result, LoadOptions.None);
                            var docs = from document in doc.Descendants("Document")
                                       select new DocumentInfo() {
                                           Name = (string)document.Attribute("Name"),
                                           Description= (string)document.Attribute("Description"),
                                           OriginalLocation = new Uri(documentsUri, (string)document.Attribute("Name")),
                                           XpsLocation = new Uri(documentsUri, (string)document.Attribute("XpsLocation"))
                                       };

                            this.documents.ItemsSource = docs;
                        }
                    }
                }
            };

            client.OpenReadAsync(documentsUri);
        }

        private void HyperlinkButton_Click(object sender, RoutedEventArgs e)
        {
            var button = (HyperlinkButton)sender;
            var document = (DocumentInfo)button.DataContext;

            if (DocumentSelected != null) {
                DocumentSelected(this, new DocumentSelectedEventArgs(document));
            }
        }
    }
}
