' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
' The path to the documents directory.
Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()

Dim doc As New Document(dataDir & Convert.ToString("Test File (doc).doc"))
For Each signature As DigitalSignature In doc.DigitalSignatures
    Console.WriteLine("*** Signature Found ***")
    Console.WriteLine("Is valid: " + signature.IsValid)
    Console.WriteLine("Reason for signing: " + signature.Comments)
    ' This property is available in MS Word documents only.
    Console.WriteLine("Time of signing: " + signature.SignTime)
    Console.WriteLine("Subject name: " + signature.Certificate.SubjectName.Name)
    Console.WriteLine("Issuer name: " + signature.Certificate.IssuerName.Name)
    Console.WriteLine()
Next
