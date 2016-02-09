Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Words

Public Class CreateDocument
    Public Shared Sub Run()
        ' ExStart:CreateDocument            
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()

        ' Initialize a Document.
        Dim doc As New Document()

        ' Use a document builder to add content to the document.
        Dim builder As New DocumentBuilder(doc)
        builder.Writeln("Hello World!")

        dataDir = dataDir & Convert.ToString("CreateDocument_out_.docx")
        ' Save the document to disk.
        doc.Save(dataDir)

        ' ExEnd:CreateDocument

        Console.WriteLine(Convert.ToString(vbLf & "Document created successfully." & vbLf & "File saved at ") & dataDir)

    End Sub
End Class
