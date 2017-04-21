Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports Aspose.Words.Replacing
Imports Aspose.Words

Public Class FindAndReplace
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_QuickStart()
        Dim fileName As String = "ReplaceSimple.doc"
        ' Open the document.
        Dim doc As New Document(dataDir & fileName)

        ' Check the text of the document
        Console.WriteLine("Original document text: " & doc.Range.Text)

        ' Replace the text in the document.
        doc.Range.Replace("_CustomerName_", "James Bond", New FindReplaceOptions(FindReplaceDirection.Forward))

        ' Check the replacement was made.
        Console.WriteLine("Document text after replace: " & doc.Range.Text)

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        ' Save the modified document.
        doc.Save(dataDir)

        Console.WriteLine(vbNewLine + "Text found and replaced successfully." + vbNewLine + "File saved at " + dataDir)
    End Sub
End Class
