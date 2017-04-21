Imports Microsoft.VisualBasic
Imports System
Imports System.Reflection
Imports System.Collections
Imports System.IO
Imports System.Text

Imports Aspose.Words.Lists
Imports Aspose.Words.Fields
Imports Aspose.Words

Public Class KeepSourceTogether
    Public Shared Sub Run()
        ' ExStart:KeepSourceTogether
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()
        Dim fileName As String = "TestFile.Destination.doc"

        Dim dstDoc As New Document(dataDir & fileName)
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        ' Set the source document to appear straight after the destination document' S content.
        srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous

        ' Iterate through all sections in the source document.
        For Each para As Paragraph In srcDoc.GetChildNodes(NodeType.Paragraph, True)
            para.ParagraphFormat.KeepWithNext = True
        Next para

        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)
        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        dstDoc.Save(dataDir)
        ' ExEnd:KeepSourceTogether
        Console.WriteLine(vbNewLine & "Document appended successfully with keeping source together." & vbNewLine & "File saved at " + dataDir)
    End Sub
End Class
