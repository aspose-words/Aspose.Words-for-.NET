Imports System.IO
Imports Aspose.Words
Imports System.Drawing
Imports Aspose.Words.Tables

Class DocumentBuilderInsertTCField
    Public Shared Sub Run()
        ' ExStart:DocumentBuilderInsertTCField
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        ' Initialize document.
        Dim doc As New Document()

        ' Create a document builder to insert content with.
        Dim builder As New DocumentBuilder(doc)

        ' Insert a TC field at the current document builder position.
        builder.InsertField("TC ""Entry Text"" \f t")

        dataDir = dataDir & Convert.ToString("DocumentBuilderInsertTCField_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderInsertTCField
        Console.WriteLine(Convert.ToString(vbLf & "TC field inserted successfully into a document." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
