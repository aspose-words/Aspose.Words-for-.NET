Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Fields
Public Class InsertMailMergeAddressBlockFieldUsingDOM
    Public Shared Sub Run()
        ' ExStart:InsertMailMergeAddressBlockFieldUsingDOM
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()
        Dim doc As New Document(dataDir & Convert.ToString("in.doc"))
        Dim builder As New DocumentBuilder(doc)

        ' Get paragraph you want to append this merge field to
        Dim para As Paragraph = DirectCast(doc.GetChildNodes(NodeType.Paragraph, True)(1), Paragraph)

        ' Move cursor to this paragraph
        builder.MoveTo(para)

        ' We want to insert a mail merge address block like this:
        ' { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\" }

        ' Create instance of FieldAddressBlock class and lets build the above field code
        Dim field As FieldAddressBlock = DirectCast(builder.InsertField(FieldType.FieldAddressBlock, False), FieldAddressBlock)

        ' { ADDRESSBLOCK \\c 1" }
        field.IncludeCountryOrRegionName = "1"

        ' { ADDRESSBLOCK \\c 1 \\d" }
        field.FormatAddressOnCountryOrRegion = True

        ' { ADDRESSBLOCK \\c 1 \\d \\e Test2 }
        field.ExcludedCountryOrRegionName = "Test2"

        ' { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 }
        field.NameAndAddressFormat = "Test3"

        ' { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\" }
        field.LanguageId = "Test 4"

        ' Finally update this merge field
        field.Update()

        dataDir = dataDir & Convert.ToString("InsertMailMergeAddressBlockFieldUsingDOM_out_.doc")
        doc.Save(dataDir)

        ' ExEnd:InsertMailMergeAddressBlockFieldUsingDOM
        Console.WriteLine(Convert.ToString(vbLf & "Mail Merge address block field using DOM inserted successfully." & vbLf & "File saved at ") & dataDir)
    End Sub

End Class
