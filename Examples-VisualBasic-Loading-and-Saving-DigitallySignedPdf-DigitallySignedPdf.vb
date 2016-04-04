' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
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
options.DigitalSignatureDetails = New PdfDigitalSignatureDetails(cert, "Test Signing", "Aspose Office", DateTime.Now)

dataDir = dataDir & Convert.ToString("Document.Signed_out_.pdf")
' Save the document as PDF with the digital signature set.
doc.Save(dataDir, options)

