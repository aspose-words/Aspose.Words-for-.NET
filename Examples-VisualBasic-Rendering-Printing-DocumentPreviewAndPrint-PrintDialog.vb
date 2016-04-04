' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim printDlg As New PrintDialog()
' Initialize the print dialog with the number of pages in the document.
printDlg.AllowSomePages = True
printDlg.PrinterSettings.MinimumPage = 1
printDlg.PrinterSettings.MaximumPage = doc.PageCount
printDlg.PrinterSettings.FromPage = 1
printDlg.PrinterSettings.ToPage = doc.PageCount
