Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.IO
Imports Aspose.Words
Imports System.Web
Public Class MailMergeUsingMustacheSyntax
    Public Shared Sub Run()
        ' ExStart:MailMergeUsingMustacheSyntax
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()
        Dim ds As New DataSet()

        ds.ReadXml(dataDir & Convert.ToString("Vendors.xml"))

        ' Open a template document.
        Dim doc As New Document(dataDir & Convert.ToString("VendorTemplate.doc"))

        doc.MailMerge.UseNonMergeFields = True

        ' Execute mail merge to fill the template with data from XML using DataSet.
        doc.MailMerge.ExecuteWithRegions(ds)
        dataDir = dataDir & Convert.ToString("MailMergeUsingMustacheSyntax_out_.docx")
        ' Save the output document.
        doc.Save(dataDir)
        ' ExEnd:MailMergeUsingMustacheSyntax
        Console.WriteLine(Convert.ToString(vbLf & "Mail merge performed with mustache syntax successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
