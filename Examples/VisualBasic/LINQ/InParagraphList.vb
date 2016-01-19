
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Aspose.Words
Imports Aspose.Words.Reporting

Namespace LINQ
    Public Class InParagraphList
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_LINQ()

            Dim fileName As String = "InParagraphList.doc"
            ' Load the template document.
            Dim doc As New Document(dataDir & fileName)

            ' Create a Reporting Engine.
            Dim engine As New ReportingEngine()

            ' Execute the build report.
            engine.BuildReport(doc, Common.GetClients(), "clients")

            dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)

            ' Save the finished document to disk.
            doc.Save(dataDir)

            Console.WriteLine(Convert.ToString(vbLf & "In-Paragraph list template document is populated with the data about clients." & vbLf & "File saved at ") & dataDir)

        End Sub
    End Class
End Namespace