Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.IO
Imports System.Reflection

Imports Aspose.Words

Public Class XMLMailMerge
    Public Shared Sub Run()
        ' ExStart:XMLMailMerge
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()

        ' Create the Dataset and read the XML.
        Dim customersDs As New DataSet()
        customersDs.ReadXml(dataDir & "Customers.xml")

        Dim fileName As String = "TestFile.doc"
        ' Open a template document.
        Dim doc As New Document(dataDir & fileName)

        ' Execute mail merge to fill the template with data from XML using DataTable.
        doc.MailMerge.Execute(customersDs.Tables("Customer"))

        dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)
        ' Save the output document.
        doc.Save(dataDir)
        ' ExEnd:XMLMailMerge
        Console.WriteLine(vbNewLine + "Mail merge performed with XML data successfully." + vbNewLine + "File saved at " + dataDir)
    End Sub
End Class
