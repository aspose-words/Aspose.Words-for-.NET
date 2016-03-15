Imports System.Collections
Imports System.IO
Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.Fields
Imports Aspose.Words.Rendering
Imports Aspose.Words.Saving
Imports Aspose.Words.Settings
Imports Aspose.Words.Tables
Public Class ChangeFieldUpdateCultureSource
    Public Shared Sub Run()
        ' ExStart:ChangeFieldUpdateCultureSource
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()
        ' We will test this functionality creating a document with two fields with date formatting            
        Dim doc As New Document()
        Dim builder As New DocumentBuilder(doc)

        ' Insert content with German locale.
        builder.Font.LocaleId = 1031
        builder.InsertField("MERGEFIELD Date1 \@ ""dddd, d MMMM yyyy""")
        builder.Write(" - ")
        builder.InsertField("MERGEFIELD Date2 \@ ""dddd, d MMMM yyyy""")
        ' Shows how to specify where the culture used for date formatting during field update and mail merge is chosen from.
        ' Set the culture used during field update to the culture used by the field.            
        doc.FieldOptions.FieldUpdateCultureSource = FieldUpdateCultureSource.FieldCode
        doc.MailMerge.Execute(New String() {"Date2"}, New Object() {New DateTime(2011, 1, 1)})
        dataDir = dataDir & Convert.ToString("Field.ChangeFieldUpdateCultureSource_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:ChangeFieldUpdateCultureSource

        Console.WriteLine(Convert.ToString(vbLf & "Culture changed successfully used in formatting fields during update." & vbLf & "File saved at ") & dataDir)
    End Sub

End Class
