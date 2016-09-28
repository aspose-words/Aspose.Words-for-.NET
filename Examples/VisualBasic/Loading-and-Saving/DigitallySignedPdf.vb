Imports System.IO
Imports Aspose.Words
Imports System
' ExStart:X509Certificates
Imports System.Security.Cryptography.X509Certificates
' ExEnd:X509Certificates
Imports Aspose.Words.Saving
Public Class DigitallySignedPdf
    Public Shared Sub Run()
        ' ExStart:DigitallySignedPdf
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()

        ' Create a simple document from scratch.
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)
        builder.Writeln("Test Signed PDF.")

        ' Load the certificate from disk.
        ' The other constructor overloads can be used to load certificates from different locations.
        Dim cert As New X509Certificate2(dataDir & Convert.ToString("signature.pfx"), "signature")

        ' Pass the certificate and details to the save options class to sign with.
        Dim options As New PdfSaveOptions()
        options.DigitalSignatureDetails = New PdfDigitalSignatureDetails()

        dataDir = dataDir & Convert.ToString("Document.Signed_out_.pdf")
        ' Save the document as PDF.
        doc.Save(dataDir)

        ' ExEnd:DigitallySignedPdf
        Console.WriteLine(Convert.ToString(vbLf & "Digitally signed PDF file created successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
