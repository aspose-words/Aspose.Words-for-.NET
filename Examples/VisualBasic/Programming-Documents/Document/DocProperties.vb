Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Properties

Public Class DocProperties
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        ' Enumerates through all built-in and custom properties in a document.
        EnumerateProperties(dataDir)
        ' Checks if a custom property with a given name exists in a document and adds few more custom document properties.
        CustomAdd(dataDir)
        ' Removes a custom document property.
        CustomRemove(dataDir)
    End Sub
    Public Shared Sub EnumerateProperties(dataDir As String)
        ' ExStart:EnumerateProperties            
        Dim fileName As String = dataDir & Convert.ToString("Properties.doc")
        Dim doc As New Document(fileName)
        Console.WriteLine("1. Document name: {0}", fileName)

        Console.WriteLine("2. Built-in Properties")
        For Each prop As DocumentProperty In doc.BuiltInDocumentProperties
            Console.WriteLine("{0} : {1}", prop.Name, prop.Value)
        Next

        Console.WriteLine("3. Custom Properties")
        For Each prop As DocumentProperty In doc.CustomDocumentProperties
            Console.WriteLine("{0} : {1}", prop.Name, prop.Value)
        Next
        ' ExEnd:EnumerateProperties
    End Sub
    Public Shared Sub CustomAdd(dataDir As String)
        ' ExStart:CustomAdd            
        Dim doc As New Document(dataDir & Convert.ToString("Properties.doc"))
        Dim props As CustomDocumentProperties = doc.CustomDocumentProperties
        If props("Authorized") Is Nothing Then
            props.Add("Authorized", True)
            props.Add("Authorized By", "John Smith")
            props.Add("Authorized Date", DateTime.Today)
            props.Add("Authorized Revision", doc.BuiltInDocumentProperties.RevisionNumber)
            props.Add("Authorized Amount", 123.45)
        End If
        ' ExEnd:CustomAdd
    End Sub
    Public Shared Sub CustomRemove(dataDir As String)
        ' ExStart:CustomRemove            
        Dim doc As New Aspose.Words.Document(dataDir & Convert.ToString("Properties.doc"))
        doc.CustomDocumentProperties.Remove("Authorized Date")
        ' ExEnd:CustomRemove
    End Sub
End Class
