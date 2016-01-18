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
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()

        Dim dstDoc As New Document(dataDir & "TestFile.Destination.doc")
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        ' Set the source document to appear straight after the destination document's content.
        srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous

        ' Iterate through all sections in the source document.
        For Each para As Paragraph In srcDoc.GetChildNodes(NodeType.Paragraph, True)
            para.ParagraphFormat.KeepWithNext = True
        Next para

        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)
        dstDoc.Save(dataDir & "TestDcc.KeepSourceTogether Out.doc")

        Console.WriteLine(vbNewLine & "Document appended successfully with keeping source together." & vbNewLine & "File saved at " + dataDir + "TestFile.KeepSourceTogether Out.docx")
    End Sub
End Class
