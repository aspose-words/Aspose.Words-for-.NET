Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.IO
Imports Aspose.Words
Imports System.Web
Imports System.Drawing
Imports Aspose.Words.MailMerging
Imports System.Data.OleDb
Public Class MailMergeImageFromBlob
    Public Shared Sub Run()
        ' ExStart:MailMergeImageFromBlob            
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_MailMergeAndReporting()
        Dim doc As New Document(dataDir & Convert.ToString("MailMerge.MergeImage.doc"))

        ' Set up the event handler for image fields.
        doc.MailMerge.FieldMergingCallback = New HandleMergeImageFieldFromBlob()

        ' Open a database connection.
        Dim connString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + RunExamples.GetDataDir_Database() + "Northwind.mdb"
        Dim conn As New OleDbConnection(connString)
        conn.Open()

        ' Open the data reader. It needs to be in the normal mode that reads all record at once.
        Dim cmd As New OleDbCommand("SELECT * FROM Employees", conn)
        Dim dataReader As IDataReader = cmd.ExecuteReader()

        ' Perform mail merge.
        doc.MailMerge.ExecuteWithRegions(dataReader, "Employees")

        ' Close the database.
        conn.Close()
        dataDir = dataDir & Convert.ToString("MailMerge.MergeImage_out_.doc")
        doc.Save(dataDir)
        ' ExEnd:MailMergeImageFromBlob
        Console.WriteLine(Convert.ToString(vbLf & "Mail merge image from blob performed successfully." & vbLf & "File saved at ") & dataDir)
    End Sub
    ' ExStart:HandleMergeImageFieldFromBlob 
    Public Class HandleMergeImageFieldFromBlob
        Implements IFieldMergingCallback
        Private Sub IFieldMergingCallback_FieldMerging(args As FieldMergingArgs) Implements IFieldMergingCallback.FieldMerging
            ' Do nothing.
        End Sub

        ''' <summary>
        ''' This is called when mail merge engine encounters Image:XXX merge field in the document.
        ''' You have a chance to return an Image object, file name or a stream that contains the image.
        ''' </summary>
        Private Sub IFieldMergingCallback_ImageFieldMerging(e As ImageFieldMergingArgs) Implements IFieldMergingCallback.ImageFieldMerging
            ' The field value is a byte array, just cast it and create a stream on it.
            Dim imageStream As New MemoryStream(DirectCast(e.FieldValue, Byte()))
            ' Now the mail merge engine will retrieve the image from the stream.
            e.ImageStream = imageStream
        End Sub
    End Class
    ' ExEnd:HandleMergeImageFieldFromBlob

End Class
