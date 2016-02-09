
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Aspose.Words
Imports Aspose.Words.Reporting

Namespace LINQ
    Public Class SingleRow
        Public Shared Sub Run()
            ' ExStart:SingleRow
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_LINQ()

            Dim fileName As String = "SingleRow.doc"
            ' Load the template document.
            Dim doc As New Document(dataDir & fileName)

            ' Load the photo and read all bytes.
            Dim imgdata As Byte() = System.IO.File.ReadAllBytes(dataDir & Convert.ToString("photo.png"))

            ' Create a Reporting Engine.
            Dim engine As New ReportingEngine()

            ' Execute the build report.
            engine.BuildReport(doc, Common.GetManager(), "manager")

            dataDir = dataDir & RunExamples.GetOutputFilePath(fileName)

            ' Save the finished document to disk.
            doc.Save(dataDir)
            ' ExEnd:SingleRow
            Console.WriteLine(Convert.ToString(vbLf & "Single row template document is populated with the data about manager." & vbLf & "File saved at ") & dataDir)

        End Sub
    End Class
End Namespace
