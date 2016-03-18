Imports System.IO
Imports Aspose.Words
Imports System.Drawing
Imports Aspose.Words.Tables
Class DocumentBuilderInsertTOC
    Public Shared Sub Run()
        ' ExStart:DocumentBuilderInsertTOC
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithDocument()
        ' Initialize document.
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' Insert a table of contents at the beginning of the document.
        builder.InsertTableOfContents("\o ""1-3"" \h \z \u")

        ' The newly inserted table of contents will be initially empty.
        ' It needs to be populated by updating the fields in the document.
        ' ExStart:UpdateFields
        doc.UpdateFields()
        ' ExEnd:UpdateFields
        dataDir = dataDir & Convert.ToString("DocumentBuilderInsertTOC_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:DocumentBuilderInsertTOC
        Console.WriteLine(Convert.ToString(vbLf & "Table of contents field inserted successfully into a document." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
