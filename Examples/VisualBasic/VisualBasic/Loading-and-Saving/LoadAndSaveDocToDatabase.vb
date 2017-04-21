Imports System.Collections
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports Aspose.Words
Public Class LoadAndSaveDocToDatabase
    Public Shared Sub Run()
        ' The path to the documents directory.
        Dim dataDir As String = RunExamples.GetDataDir_LoadingAndSaving()
        Dim fileName As String = "TestFile.doc"
        ' Open the document.
        Dim doc As New Document(dataDir & fileName)
        ' ExStart:OpenDatabaseConnection 
        Dim dbName As String = ""
        ' Open a database connection.
        Dim connString As String = Convert.ToString("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + RunExamples.GetDataDir_Database()) & dbName
        Dim mConnection As New OleDbConnection(connString)
        mConnection.Open()
        ' ExEnd:OpenDatabaseConnection
        ' ExStart:OpenRetrieveAndDelete 
        ' Store the document to the database.
        StoreToDatabase(doc, mConnection)
        ' Read the document from the database and store the file to disk.
        Dim dbDoc As Document = ReadFromDatabase(fileName, mConnection)

        ' Save the retrieved document to disk.
        Dim newFileName As String = Path.GetFileNameWithoutExtension(fileName) + " from DB" + Path.GetExtension(fileName)
        dbDoc.Save(dataDir & newFileName)

        ' Delete the document from the database.
        DeleteFromDatabase(fileName, mConnection)

        ' Close the connection to the database.
        mConnection.Close()
        ' ExEnd:OpenRetrieveAndDelete 

    End Sub
    ' ExStart:StoreToDatabase 
    Public Shared Sub StoreToDatabase(doc As Document, mConnection As OleDbConnection)
        ' Save the document to a MemoryStream object.
        Dim stream As New MemoryStream()
        doc.Save(stream, SaveFormat.Doc)

        ' Get the filename from the document.
        Dim fileName As String = Path.GetFileName(doc.OriginalFileName)

        ' Create the SQL command.
        Dim commandString As String = (Convert.ToString("INSERT INTO Documents (FileName, FileContent) VALUES('") & fileName) + "', @Doc)"
        Dim command As New OleDbCommand(commandString, mConnection)

        ' Add the @Doc parameter.
        command.Parameters.AddWithValue("Doc", stream.ToArray())

        ' Write the document to the database.
        command.ExecuteNonQuery()
    End Sub
    ' ExEnd:StoreToDatabase
    ' ExStart:ReadFromDatabase 
    Public Shared Function ReadFromDatabase(fileName As String, mConnection As OleDbConnection) As Document
        ' Create the SQL command.
        Dim commandString As String = (Convert.ToString("SELECT * FROM Documents WHERE FileName='") & fileName) + "'"
        Dim command As New OleDbCommand(commandString, mConnection)

        ' Create the data adapter.
        Dim adapter As New OleDbDataAdapter(command)

        ' Fill the results from the database into a DataTable.
        Dim dataTable As New DataTable()
        adapter.Fill(dataTable)

        ' Check there was a matching record found from the database and throw an exception if no record was found.
        If dataTable.Rows.Count = 0 Then
            Throw New ArgumentException(String.Format("Could not find any record matching the document ""{0}"" in the database.", fileName))
        End If

        ' The document is stored in byte form in the FileContent column.
        ' Retrieve these bytes of the first matching record to a new buffer.
        Dim buffer As Byte() = DirectCast(dataTable.Rows(0)("FileContent"), Byte())

        ' Wrap the bytes from the buffer into a new MemoryStream object.
        Dim newStream As New MemoryStream(buffer)

        ' Read the document from the stream.
        Dim doc As New Document(newStream)

        ' Return the retrieved document.
        Return doc
    End Function
    ' ExEnd:ReadFromDatabase
    ' ExStart:DeleteFromDatabase 
    Public Shared Sub DeleteFromDatabase(fileName As String, mConnection As OleDbConnection)
        ' Create the SQL command.
        Dim commandString As String = (Convert.ToString("DELETE * FROM Documents WHERE FileName='") & fileName) + "'"
        Dim command As New OleDbCommand(commandString, mConnection)

        ' Delete the record.
        command.ExecuteNonQuery()
    End Sub
    ' ExEnd:DeleteFromDatabase
End Class

