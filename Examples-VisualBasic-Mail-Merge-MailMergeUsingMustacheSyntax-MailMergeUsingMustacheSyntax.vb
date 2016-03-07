' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
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
