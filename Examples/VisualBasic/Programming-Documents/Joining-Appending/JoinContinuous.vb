Imports Microsoft.VisualBasic
Imports System
Imports System.Reflection
Imports System.Collections
Imports System.IO
Imports System.Text

Imports Aspose.Words.Lists
Imports Aspose.Words.Fields
Imports Aspose.Words

Public Class JoinContinuous
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()

        Dim dstDoc As New Document(dataDir & "TestFile.Destination.doc")
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        ' Make the document appear straight after the destination documents content.
        srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous

        ' Append the source document using the original styles found in the source document.
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)
        dstDoc.Save(dataDir & "TestFile.JoinContinuous Out.doc")

        Console.WriteLine(vbNewLine & "Document appended successfully with join continous option." & vbNewLine & "File saved at " + dataDir + "TestFile.JoinContinuous Out.docx")
    End Sub
End Class
