Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Fields
Public Class InsertMergeFieldUsingDOM
    Public Shared Sub Run()
        ' ExStart:InsertMergeFieldUsingDOM
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()
        Dim doc As New Document(dataDir & Convert.ToString("in.doc"))
        Dim builder As New DocumentBuilder(doc)

        ' Get paragraph you want to append this merge field to
        Dim para As Paragraph = DirectCast(doc.GetChildNodes(NodeType.Paragraph, True)(1), Paragraph)

        ' Move cursor to this paragraph
        builder.MoveTo(para)

        ' We want to insert a merge field like this:
        ' { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }

        ' Create instance of FieldMergeField class and lets build the above field code
        Dim field As FieldMergeField = DirectCast(builder.InsertField(FieldType.FieldMergeField, False), FieldMergeField)

        ' { " MERGEFIELD Test1" }
        field.FieldName = "Test1"

        ' { " MERGEFIELD Test1 \\b Test2" }
        field.TextBefore = "Test2"

        ' { " MERGEFIELD Test1 \\b Test2 \\f Test3 }
        field.TextAfter = "Test3"

        ' { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m" }
        field.IsMapped = True

        ' { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }
        field.IsVerticalFormatting = True

        ' Finally update this merge field
        field.Update()

        dataDir = dataDir & Convert.ToString("InsertMergeFieldUsingDOM_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:InsertMergeFieldUsingDOM
        Console.WriteLine(Convert.ToString(vbLf & "Merge field using DOM inserted successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
