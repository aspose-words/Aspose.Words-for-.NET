
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports Aspose.Words
Imports Aspose.Words.Reporting

Namespace LINQ
    Public Class InTableMasterDetail
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_LINQ()

            ' Load the template document.
            Dim doc As New Document(dataDir & Convert.ToString("InTableMasterDetail.doc"))

            ' Create a Reporting Engine.
            Dim engine As New ReportingEngine()

            ' Execute the build report.
            engine.BuildReport(doc, Common.GetManagers(), "managers")

            dataDir = dataDir & Convert.ToString("InTableMasterDetail Out.doc")

            ' Save the finished document to disk.
            doc.Save(dataDir)

            Console.WriteLine(Convert.ToString(vbLf & "In-Table master detail template document is populated with the data about managers and it's contracts." & vbLf & "File saved at ") & dataDir)

        End Sub
    End Class
End Namespace