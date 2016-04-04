// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
//ExId:DocumentPreviewAndPrint_ActivePrintPreviewDialog_Creation
//ExSummary:Creates an overridden version of the .NET Print Preview dialog to preview the document.
ActivePrintPreviewDialog previewDlg = new ActivePrintPreviewDialog();

// Pass the Aspose.Words print document to the Print Preview dialog.
previewDlg.Document = awPrintDoc;
// Specify additional parameters of the Print Preview dialog.
previewDlg.ShowInTaskbar = true;
previewDlg.MinimizeBox = true;
previewDlg.PrintPreviewControl.Zoom = 1;
previewDlg.Document.DocumentName = "TestName.doc";
previewDlg.WindowState = FormWindowState.Maximized;
// Show the appropriately configured Print Preview dialog.
previewDlg.ShowDialog();
