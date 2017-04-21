Imports Microsoft.VisualBasic
Imports Aspose.Words
Public Class ProtectDocument
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        dataDir = dataDir & Convert.ToString("ProtectDocument.doc")

        Protect(dataDir)
        UnProtect(dataDir)
        GetProtectionType(dataDir)
    End Sub
    ''' <summary>
    ''' Shows how to protect document
    ''' </summary>
    ''' <param name="inputFileName">input file name with complete path.</param>        
    Public Shared Sub Protect(inputFileName As String)
        ' ExStart:ProtectDocument
        Dim doc As New Document(inputFileName)
        doc.Protect(ProtectionType.AllowOnlyFormFields, "password")
        ' ExEnd:ProtectDocument
        Console.WriteLine(vbLf & "Document protected successfully.")

    End Sub
    ''' <summary>
    ''' Shows how to unprotect document
    ''' </summary>
    ''' <param name="inputFileName">input file name with complete path.</param>        
    Public Shared Sub UnProtect(inputFileName As String)
        ' ExStart:UnProtectDocument
        Dim doc As New Document(inputFileName)
        doc.Unprotect()
        ' ExEnd:UnProtectDocument
        Console.WriteLine(vbLf & "Document unprotected successfully.")
    End Sub
    ''' <summary>
    ''' Shows how to get protection type
    ''' </summary>
    ''' <param name="inputFileName">input file name with complete path.</param>        
    Public Shared Sub GetProtectionType(inputFileName As String)
        ' ExStart:GetProtectionType
        Dim doc As New Document(inputFileName)
        Dim protectionType As ProtectionType = doc.ProtectionType
        ' ExEnd:GetProtectionType
        Console.WriteLine(vbLf & "Document protection type is " + protectionType.ToString())
    End Sub
End Class
