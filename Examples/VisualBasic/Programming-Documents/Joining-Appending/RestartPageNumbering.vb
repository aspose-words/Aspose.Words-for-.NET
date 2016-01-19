Imports Microsoft.VisualBasic
Imports System
Imports System.Reflection
Imports System.Collections
Imports System.IO
Imports System.Text

Imports Aspose.Words.Lists
Imports Aspose.Words.Fields
Imports Aspose.Words

Public Class RestartPageNumbering
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()

        Dim dstDoc As New Document(dataDir & "TestFile.Destination.doc")
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        ' Set the appended document to appear on the next page.
        srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage
        ' Restart the page numbering for the document to be appended.
        srcDoc.FirstSection.PageSetup.RestartPageNumbering = True

        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)
        dstDoc.Save(dataDir & "TestFile.RestartPageNumbering Out.doc")

        Console.WriteLine(vbNewLine & "Document appended successfully with restart page numbering option." & vbNewLine & "File saved at " + dataDir + "TestFile.RestartPageNumbering Out.docx")
    End Sub
End Class
