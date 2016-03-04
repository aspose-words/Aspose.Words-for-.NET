// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
//ExId:DocumentPreviewAndPrint_AsposeWordsPrintDocument_Creation
//ExSummary:Creates a special Aspose.Words implementation of the .NET PrintDocument class.
// Pass the printer settings from the dialog to the print document.
AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
awPrintDoc.PrinterSettings = printDlg.PrinterSettings;
