
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Aspose.Words
Imports Aspose.Words.Reporting

Namespace LINQ
    Public Class InTableWithFilteringGroupingSorting
        Public Shared Sub Run()
            ' ExStart:InTableWithFilteringGroupingSorting
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_LINQ()

            Dim fileName As String = "InTableWithFilteringGroupingSorting.doc"
            ' Load the template document.
            Dim doc As New Document(dataDir & fileName)

            ' Create a Reporting Engine.
            Dim engine As New ReportingEngine()

            ' Execute the build report.
            engine.BuildReport(doc, Common.GetContracts(), "contracts")

            dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)

            ' Save the finished document to disk.
            doc.Save(dataDir)
            ' ExEnd:InTableWithFilteringGroupingSorting
            Console.WriteLine(Convert.ToString(vbLf & "In-Table filtering, grouping and sorting template document is populated with the data about managers." & vbLf & "File saved at ") & dataDir)

        End Sub
    End Class
End Namespace