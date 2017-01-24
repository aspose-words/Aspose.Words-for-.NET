Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Words
Public Class SendToClientBrowser
    Public Shared Sub Run()
        ' ExStart:SendToClientBrowser
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()

        Dim doc As New Document(dataDir & Convert.ToString("Document.doc"))

        dataDir = dataDir & Convert.ToString("Report_out.doc")
        ' If this method overload is causing a compiler error then you are using the Client Profile DLL whereas 
        ' The Aspose.Words .NET 2.0 DLL must be used instead.
        doc.Save(dataDir)
        ' ExEnd:SendToClientBrowser
        Console.WriteLine(Convert.ToString(vbLf & "Document send to client browser successfully." & vbLf & "File saved at ") & dataDir)
    End Sub

End Class
