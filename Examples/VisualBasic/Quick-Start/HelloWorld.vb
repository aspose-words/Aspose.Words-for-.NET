Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Words

Public Class HelloWorld
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_QuickStart()

        ' Create a blank document.
        Dim doc As New Document()

        ' DocumentBuilder provides members to easily add content to a document.
        Dim builder As New DocumentBuilder(doc)

        ' Write a new paragraph in the document with the text "Hello World!"
        builder.Writeln("Hello World!")

        ' Save the document in DOCX format. The format to save as is inferred from the extension of the file name.
        ' Aspose.Words supports saving any document in many more formats.
        dataDir = dataDir & "HelloWorld_out_.docx"
        doc.Save(dataDir)

        Console.WriteLine(vbNewLine + "New document created successfully." + vbNewLine + "File saved at " + dataDir)
    End Sub
End Class
