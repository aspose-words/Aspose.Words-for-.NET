Imports System.Collections
Imports System.IO
Imports Aspose.Words
Imports System.Threading
Imports System.Globalization
Public Class ChangeLocale
    Public Shared Sub Run()
        ' ExStart:ChangeLocale
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_WorkingWithFields()

        ' Create a blank document.
        Dim doc As New Document()
        Dim b As New DocumentBuilder(doc)
        b.InsertField("MERGEFIELD Date")

        ' Store the current culture so it can be set back once mail merge is complete.
        Dim currentCulture As CultureInfo = Thread.CurrentThread.CurrentCulture
        ' Set to German language so dates and numbers are formatted using this culture during mail merge.
        Thread.CurrentThread.CurrentCulture = New CultureInfo("de-DE")

        ' Execute mail merge.
        doc.MailMerge.Execute(New String() {"Date"}, New Object() {DateTime.Now})

        ' Restore the original culture.
        Thread.CurrentThread.CurrentCulture = currentCulture
        doc.Save(dataDir & Convert.ToString("Field.ChangeLocale_out.doc"))
        ' ExEnd:ChangeLocale

        Console.WriteLine(Convert.ToString(vbLf & "Culture changed successfully used in formatting fields during update." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
