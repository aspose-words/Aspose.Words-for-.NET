
Imports Microsoft.VisualBasic
Imports System
Imports System.Reflection
Imports System.Collections
Imports System.IO
Imports System.Text

Imports Aspose.Words.Lists
Imports Aspose.Words.Fields
Imports Aspose.Words

Public Class UseDestinationStyles
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_JoiningAndAppending()
        Dim fileName As String = "TestFile.Destination.doc"

        Dim dstDoc As New Document(dataDir & fileName)
        Dim srcDoc As New Document(dataDir & "TestFile.Source.doc")

        ' Append the source document using the styles of the destination document.
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles)

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        ' Save the joined document to disk.
        dstDoc.Save(dataDir)

        Console.WriteLine(vbNewLine & "Document appended successfully with use destination styles option." & vbNewLine & "File saved at " + dataDir)
    End Sub
End Class
