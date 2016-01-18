Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.IO
Imports System.Reflection
Imports System.Diagnostics

Imports Aspose.Words
Imports System.Collections

Public Class NestedMailMerge
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()

        ' Create the Dataset and read the XML.
        Dim pizzaDs As New DataSet()

        ' Note: The Datatable.TableNames and the DataSet.Relations are defined implicitly by .NET through ReadXml.
        ' To see examples of how to set up relations manually check the corresponding documentation of this sample
        pizzaDs.ReadXml(dataDir & "CustomerData.xml")

        ' Open the template document.
        Dim doc As New Document(dataDir & "Invoice Template.doc")

        ' Trim trailing and leading whitespaces mail merge values
        doc.MailMerge.TrimWhitespaces = False

        ' Execute the nested mail merge with regions
        doc.MailMerge.ExecuteWithRegions(pizzaDs)

        ' Save the output to file
        doc.Save(dataDir & "Invoice Out.doc")

        Debug.Assert(doc.MailMerge.GetFieldNames().Length = 0, "There was a problem with mail merge")

        Console.WriteLine(vbNewLine + "Mail merge performed with nested data successfully." + vbNewLine + "File saved at " + dataDir + "Invoice Out.doc")
    End Sub
End Class
