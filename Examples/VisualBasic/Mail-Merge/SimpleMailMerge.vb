Imports Aspose.Words
Imports Aspose.Words.MailMerging
Imports System.Collections.Generic
Imports System.Web

Class SimpleMailMerge
    Public Shared Sub Run()
        ' ExStart:SimpleMailMerge
        Dim Response As HttpResponse = Nothing
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()
        ' Open an existing document.
        Dim doc As New Document(dataDir & Convert.ToString("MailMerge.ExecuteArray.doc"))

        doc.MailMerge.UseNonMergeFields = True

        ' Fill the fields in the document with user data.
        doc.MailMerge.Execute(New String() {"FullName", "Company", "Address", "Address2", "City"}, New Object() {"James Bond", "MI5 Headquarters", "Milbank", "", "London"})

        dataDir = dataDir & Convert.ToString("MailMerge.ExecuteArray_out_.doc")
        ' Send the document in Word format to the client browser with an option to save to disk or open inside the current browser.
        doc.Save(Response, dataDir, ContentDisposition.Inline, Nothing)
        ' ExEnd:SimpleMailMerge
        Console.WriteLine(Convert.ToString(vbLf & "Simple Mail merge performed with array data successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
End Class
