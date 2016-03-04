' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim previewDlg As New ActivePrintPreviewDialog()

' Pass the Aspose.Words print document to the Print Preview dialog.
previewDlg.Document = awPrintDoc
' Specify additional parameters of the Print Preview dialog.
previewDlg.ShowInTaskbar = True
previewDlg.MinimizeBox = True
previewDlg.PrintPreviewControl.Zoom = 1
previewDlg.Document.DocumentName = "TestName.doc"
previewDlg.WindowState = FormWindowState.Maximized
' Show the appropriately configured Print Preview dialog.
previewDlg.ShowDialog()
