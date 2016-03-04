// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
//ExId:DocumentPreviewAndPrint_PrintDialog_Creation
//ExSummary:Creates the print dialog.
PrintDialog printDlg = new PrintDialog();
// Initialize the print dialog with the number of pages in the document.
printDlg.AllowSomePages = true;
printDlg.PrinterSettings.MinimumPage = 1;
printDlg.PrinterSettings.MaximumPage = doc.PageCount;
printDlg.PrinterSettings.FromPage = 1;
printDlg.PrinterSettings.ToPage = doc.PageCount;
