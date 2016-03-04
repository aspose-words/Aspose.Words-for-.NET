// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
//ExId:DocumentPreviewAndPrint_PrintDialog_Check_Result
//ExSummary:Check if the user accepted the print settings and proceed to preview the document.
if (!printDlg.ShowDialog().Equals(DialogResult.OK))
    return;
