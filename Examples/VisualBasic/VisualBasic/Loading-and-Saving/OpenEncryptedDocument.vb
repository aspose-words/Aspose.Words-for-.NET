Imports System.Collections
Imports System.IO
Imports Aspose.Words

Public Class OpenEncryptedDocument
   Public Shared Sub Run()
        ' ExStart:OpenEncryptedDocument      
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()

        ' Loads encrypted document.
        Dim doc As New Document(dataDir & Convert.ToString("LoadEncrypted.docx"), New LoadOptions("aspose"))

        ' ExEnd:OpenEncryptedDocument

        Console.WriteLine(vbLf & "Encrypted document loaded successfully.")

    End Sub
End Class
