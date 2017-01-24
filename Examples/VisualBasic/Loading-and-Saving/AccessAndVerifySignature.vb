Imports Microsoft.VisualBasic
Imports System
Imports Aspose.Words

Public Class AccessAndVerifySignature
   Public Shared Sub Run()
        ' ExStart:AccessAndVerifySignature            
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()

        Dim doc As New Document(dataDir & Convert.ToString("Test File (doc).doc"))
        For Each signature As DigitalSignature In doc.DigitalSignatures
            Console.WriteLine("*** Signature Found ***")
            Console.WriteLine("Is valid: " + signature.IsValid.ToString())
            Console.WriteLine("Reason for signing: " + signature.Comments.ToString())
            ' This property is available in MS Word documents only.
            Console.WriteLine("Time of signing: " + signature.SignTime.ToString())
            Console.WriteLine("Subject name: " + signature.CertificateHolder.Certificate.SubjectName.Name)
            Console.WriteLine("Issuer name: " + signature.CertificateHolder.Certificate.IssuerName.Name)
            Console.WriteLine()
        Next
        ' ExEnd:AccessAndVerifySignature
    End Sub
End Class
