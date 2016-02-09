Imports Microsoft.VisualBasic
Imports System.IO
Imports Aspose.Words
Public Class ConvertDocumentToByte
    Public Shared Sub Run()
        ' ExStart:ConvertDocumentToByte
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()

        ' Load the document from disk.
        Dim doc As New Document(dataDir & Convert.ToString("Test File (doc).doc"))

        ' Create a new memory stream.
        Dim outStream As New MemoryStream()
        ' Save the document to stream.
        doc.Save(outStream, SaveFormat.Docx)

        ' Convert the document to byte form.
        Dim docBytes As Byte() = outStream.ToArray()

        ' The bytes are now ready to be stored/transmitted.

        ' Now reverse the steps to load the bytes back into a document object.
        Dim inStream As New MemoryStream(docBytes)

        ' Load the stream into a new document object.
        Dim loadDoc As New Document(inStream)
        ' ExEnd:ConvertDocumentToByte

        Console.WriteLine(vbLf & "Document converted to byte array successfully.")
    End Sub
End Class
