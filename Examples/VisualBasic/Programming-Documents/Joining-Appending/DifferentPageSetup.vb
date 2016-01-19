Imports Microsoft.VisualBasic
Imports System
Imports System.Reflection
Imports System.Collections
Imports System.IO
Imports System.Text

Imports Aspose.Words.Lists
Imports Aspose.Words.Fields
Imports Aspose.Words

Public Class DifferentPageSetup
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()

        Dim dstDoc As New Document(dataDir & "TestFile.Destination.doc")
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        ' Set the source document to continue straight after the end of the destination document.
        ' If some page setup settings are different then this may not work and the source document will appear 
        ' on a new page.
        srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous

        ' To ensure this does not happen when the source document has different page setup settings make sure the
        ' settings are identical between the last section of the destination document.
        ' If there are further continuous sections that follow on in the source document then this will need to be 
        ' repeated for those sections as well.
        srcDoc.FirstSection.PageSetup.PageWidth = dstDoc.LastSection.PageSetup.PageWidth
        srcDoc.FirstSection.PageSetup.PageHeight = dstDoc.LastSection.PageSetup.PageHeight
        srcDoc.FirstSection.PageSetup.Orientation = dstDoc.LastSection.PageSetup.Orientation

        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)
        dstDoc.Save(dataDir & "TestFile.DifferentPageSetup Out.doc")

        Console.WriteLine(vbNewLine & "Document appended successfully with different page setup." & vbNewLine & "File saved at " + dataDir + "TestFile.DifferentPageSetup Out.docx")
    End Sub
End Class
