Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Words

Public Class SimpleMailMerge
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_QuickStart()

        Dim doc As New Document(dataDir & "Template.doc")

        ' Fill the fields in the document with user data.
        doc.MailMerge.Execute(New String() {"FullName", "Company", "Address", "Address2", "City"}, New Object() {"James Bond", "MI5 Headquarters", "Milbank", "", "London"})

        ' Saves the document to disk.
        doc.Save(dataDir & "MailMerge Result Out.docx")

        Console.WriteLine(vbNewLine + "Mail merge performed successfully." + vbNewLine + "File saved at " + dataDir + "MailMerge Result Out.docx")
    End Sub
End Class
