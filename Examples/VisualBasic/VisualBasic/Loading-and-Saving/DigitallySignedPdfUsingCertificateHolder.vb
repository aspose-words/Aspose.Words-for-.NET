Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Saving
Class DigitallySignedPdfUsingCertificateHolder
    Public Shared Sub Run()
        ' ExStart:DigitallySignedPdfUsingCertificateHolder
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()

        ' Create a simple document from scratch.
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)
        builder.Writeln("Test Signed PDF.")

        Dim options As New PdfSaveOptions()
        options.DigitalSignatureDetails = New PdfDigitalSignatureDetails(CertificateHolder.Create(dataDir & Convert.ToString("CioSrv1.pfx"), "cinD96..arellA"), "reason", "location", DateTime.Now)
        doc.Save(dataDir & Convert.ToString("DigitallySignedPdfUsingCertificateHolder.Signed_out.pdf"), options)
        ' ExEnd:DigitallySignedPdfUsingCertificateHolder
        Console.WriteLine(Convert.ToString(vbLf & "Digitally signed PDF file created successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class