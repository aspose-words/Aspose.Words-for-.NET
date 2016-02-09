
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Aspose.Words
Imports Aspose.Words.Reporting

Namespace LINQ
    Public Class CommonMasterDetail
        Public Shared Sub Run()
            ' ExStart:CommonMasterDetail
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_LINQ()

            Dim fileName As String = "CommonMasterDetail.doc"
            ' Load the template document.
            Dim doc As New Document(dataDir & fileName)

            ' Create a Reporting Engine.
            Dim engine As New ReportingEngine()

            ' Execute the build report.
            engine.BuildReport(doc, Common.GetManagers(), "managers")

            dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)

            ' Save the finished document to disk.
            doc.Save(dataDir)
            ' ExEnd:CommonMasterDetail
            Console.WriteLine(Convert.ToString(vbLf & "Common master detail template document is populated with the data about managers and it's contracts." & vbLf & "File saved at ") & dataDir)

        End Sub
    End Class
End Namespace