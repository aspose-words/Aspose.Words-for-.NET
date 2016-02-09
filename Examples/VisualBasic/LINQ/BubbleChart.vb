
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Aspose.Words
Imports Aspose.Words.Reporting

Namespace LINQ
    Public Class BubbleChart
        Public Shared Sub Run()
            ' ExStart:BubbleChart
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_LINQ()

            Dim fileName As String = "BubbleChart.docx"
            ' Load the template document.
            Dim doc As New Document(dataDir & fileName)

            ' Create a Reporting Engine.
            Dim engine As New ReportingEngine()

            ' Execute the build report.
            engine.BuildReport(doc, Common.GetContracts(), "contracts")

            dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)

            ' Save the finished document to disk.
            doc.Save(dataDir)
            ' ExEnd:BubbleChart
            Console.WriteLine(Convert.ToString(vbLf & "Bubble chart template document is populated with the data about contracts." & vbLf & "File saved at ") & dataDir)

        End Sub
    End Class
End Namespace
