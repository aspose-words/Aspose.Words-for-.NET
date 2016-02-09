

Imports System.Collections
Imports System.IO

Imports Aspose.Words
Imports Aspose.Words.Tables
Imports System.Diagnostics
Imports Aspose.Words.MailMerging
Imports Aspose.Words.Saving
Imports System.Text

Public Class LoadTxt
    Public Shared Sub Run()
        ' ExStart:LoadTxt
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()

        ' The encoding of the text file is automatically detected.
        Dim doc As New Document(dataDir & Convert.ToString("LoadTxt.txt"))

        dataDir = dataDir & "LoadTxt_out_.docx"
        ' Save as any Aspose.Words supported format, such as DOCX.
        doc.Save(dataDir)
        ' ExEnd:LoadTxt

        Console.WriteLine(vbNewLine + "Text document loaded successfully." + vbNewLine + "File saved at " + dataDir)
    End Sub
End Class
