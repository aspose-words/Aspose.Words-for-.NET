Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Fields
Public Class InsertAdvanceFieldWithOutDocumentBuilder
    Public Shared Sub Run()
        ' ExStart:InsertAdvanceFieldWithOutDocumentBuilder
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()
        Dim doc As New Document(dataDir & Convert.ToString("in.doc"))
        ' Get paragraph you want to append this Advance field to
        Dim para As Paragraph = DirectCast(doc.GetChildNodes(NodeType.Paragraph, True)(1), Paragraph)

        ' We want to insert an Advance field like this:
        ' { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 \\y 100 }

        ' Create instance of FieldAdvance class and lets build the above field code
        Dim field As FieldAdvance = DirectCast(para.AppendField(FieldType.FieldAdvance, False), FieldAdvance)


        ' { ADVANCE \\d 10 " }
        field.DownOffset = "10"

        ' { ADVANCE \\d 10 \\l 10 }
        field.LeftOffset = "10"

        ' { ADVANCE \\d 10 \\l 10 \\r -3.3 }
        field.RightOffset = "-3.3"

        ' { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 }
        field.UpOffset = "0"

        ' { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 }
        field.HorizontalPosition = "100"

        ' { ADVANCE \\d 10 \\l 10 \\r -3.3 \\u 0 \\x 100 \\y 100 }
        field.VerticalPosition = "100"

        ' Finally update this Advance field
        field.Update()

        dataDir = dataDir & Convert.ToString("InsertAdvanceFieldWithOutDocumentBuilder_out_.doc")
        doc.Save(dataDir)

        ' ExEnd:InsertAdvanceFieldWithOutDocumentBuilder
        Console.WriteLine(Convert.ToString(vbLf & "Advance field without using document builder inserted successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
