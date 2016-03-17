Imports Microsoft.VisualBasic
Imports System
Imports System.Reflection
Imports System.Collections
Imports System.IO
Imports System.Text

Imports Aspose.Words.Lists
Imports Aspose.Words.Fields
Imports Aspose.Words

Public Class UpdatePageLayout
    Public Shared Sub Run()
        ' ExStart:UpdatePageLayout
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()
        Dim fileName As String = "TestFile.Destination.doc"

        Dim dstDoc As New Document(dataDir & fileName)
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        ' If the destination document is rendered to PDF, image etc or UpdatePageLayout is called before the source document 
        ' is appended then any changes made after will not be reflected in the rendered output.
        dstDoc.UpdatePageLayout()

        ' Join the documents.
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting)

        ' For the changes to be updated to rendered output, UpdatePageLayout must be called again.
        ' If not called again the appended document will not appear in the output of the next rendering.
        dstDoc.UpdatePageLayout()

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)

        ' Save the joined document to PDF.
        dstDoc.Save(dataDir)
        ' ExEnd:UpdatePageLayout
        Console.WriteLine(vbNewLine & "Document appended successfully with updated page layout." & vbNewLine & "File saved at " + dataDir)
    End Sub
End Class
