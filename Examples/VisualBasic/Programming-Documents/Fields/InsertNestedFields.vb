Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Words
Imports Aspose.Words.Fields

Public Class InsertNestedFields
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()

        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' Insert a few page breaks (just for testing)
        For i As Integer = 0 To 4
            builder.InsertBreak(BreakType.PageBreak)
        Next i

        ' Move the DocumentBuilder cursor into the primary footer.
        builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary)

        ' We want to insert a field like this:
        ' { IF {PAGE} <> {NUMPAGES} "See Next Page" "Last Page" }
        Dim field As Field = builder.InsertField("IF ")
        builder.MoveTo(field.Separator)
        builder.InsertField("PAGE")
        builder.Write(" <> ")
        builder.InsertField("NUMPAGES")
        builder.Write(" ""See Next Page"" ""Last Page"" ")

        ' Finally update the outer field to recalcaluate the final value. Doing this will automatically update
        ' the inner fields at the same time.
        field.Update()
        dataDir = dataDir & "InsertNestedFields_out_.docx"
        doc.Save(dataDir)

        Console.WriteLine(vbNewLine & "Inserted nested fields in the document successfully." & vbNewLine & "File saved at " + dataDir)
    End Sub
End Class
