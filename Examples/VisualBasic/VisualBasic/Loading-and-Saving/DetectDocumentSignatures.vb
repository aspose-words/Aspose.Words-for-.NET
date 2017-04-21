Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Words
Public Class DetectDocumentSignatures
    Public Shared Sub Run()
        ' ExStart:DetectDocumentSignatures
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()

        ' The path to the document which is to be processed.
        Dim filePath As String = dataDir & Convert.ToString("Document.Signed.docx")

        Dim info As FileFormatInfo = FileFormatUtil.DetectFileFormat(filePath)
        If info.HasDigitalSignature Then
            Console.WriteLine(String.Format("Document {0} has digital signatures, they will be lost if you open/save this document with Aspose.Words.", Path.GetFileName(filePath)))
        End If
        ' ExEnd:DetectDocumentSignatures            
    End Sub
End Class
