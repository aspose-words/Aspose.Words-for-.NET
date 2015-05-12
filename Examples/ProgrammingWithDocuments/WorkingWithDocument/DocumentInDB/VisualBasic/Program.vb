'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports System.Reflection

Imports Aspose.Words

Namespace DocumentInDBExample
	''' <summary>
	''' This project is set to target the x86 platform because there is no 64-bit driver available 
	''' for the Access database used in this sample.
	''' </summary>
	Friend Class Program
		Public Shared Sub Main()
			' Sample infrastructure.
			Dim exeDir As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar
			Dim dataDir As String = New Uri(New Uri(exeDir), "../../Data/").LocalPath
			Dim dbName As String = dataDir & "DocDB.mdb"
			Const fileName As String = "TestFile.doc"

			'ExStart
			'ExId:DocumentInDB_DatabaseHelpers
			'ExSummary:Shows how to setup a connection to a database and execute commands.
			' Create a connection to the database.
			mConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbName)

			' Open the database connection.
			mConnection.Open()
			'ExEnd

			' Open the document.
			Dim doc As New Document(dataDir & fileName)

			'ExStart
			'ExId:DocumentInDB_Main
			'ExSummary:Stores the document to a database, then reads the same document back again, and finally deletes the record containing the document from the database.
			' Store the document to the database.
			StoreToDatabase(doc)
			' Read the document from the database and store the file to disk.
			Dim dbDoc As Document = ReadFromDatabase(fileName, dataDir)

			' Save the retrieved document to disk.
			Dim newFileName As String = Path.GetFileNameWithoutExtension(fileName) & " from DB" & Path.GetExtension(fileName)
			dbDoc.Save(dataDir & newFileName)

			' Delete the document from the database.
			DeleteFromDatabase(fileName)

			' Close the connection to the database.
			mConnection.Close()
			'ExEnd
		End Sub

		''' <summary>
		''' Stores a document object to the specified database.
		''' </summary>
		''' <param name="dbName">The name of the database file.</param>
		''' <param name="doc">The source document.</param>
		''' 
		'ExStart
		'ExId:DocumentInDB_StoreToDB
		'ExSummary:Stores the document to the specified database.
		Public Shared Sub StoreToDatabase(ByVal doc As Document)
			' Save the document to a MemoryStream object.
			Dim stream As New MemoryStream()
			doc.Save(stream, SaveFormat.Doc)

			' Get the filename from the document.
			Dim fileName As String = Path.GetFileName(doc.OriginalFileName)

			' Create the SQL command.
			Dim commandString As String = "INSERT INTO Documents (FileName, FileContent) VALUES('" & fileName & "', @Doc)"
			Dim command As New OleDbCommand(commandString, mConnection)

			' Add the @Doc parameter.
			command.Parameters.AddWithValue("Doc", stream.ToArray())

			' Write the document to the database.
			command.ExecuteNonQuery()

		End Sub
		'ExEnd

		''' <summary>
		''' Retreives a document from the specified database and saves it to disk.
		''' </summary>
		''' <param name="dbName">The name of the database file.</param>
		''' <param name="fileName">The name of the document file.</param>
		''' <param name="path">The path to the directory where to extract the document.</param>
		'ExStart
		'ExId:DocumentInDB_ReadFromDB
		'ExSummary:Retrieves and returns the document from the specified database using the filename as a key to fetch the document.
		Public Shared Function ReadFromDatabase(ByVal fileName As String, ByVal path As String) As Document
			' Create the SQL command.
			Dim commandString As String = "SELECT * FROM Documents WHERE FileName='" & fileName & "'"
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
			Dim buffer() As Byte = CType(dataTable.Rows(0)("FileContent"), Byte())

			' Wrap the bytes from the buffer into a new MemoryStream object.
			Dim newStream As New MemoryStream(buffer)

			' Read the document from the stream.
			Dim doc As New Document(newStream)

			' Return the retrieved document.
			Return doc

		End Function
		'ExEnd

		''' <summary>
		''' Deletes the records containing the specified document name from the database.
		''' </summary>
		''' <param name="dbName">The name of the database file.</param>
		''' <param name="fileName">The name of the document file.</param>
		'ExStart
		'ExId:DocumentInDB_DeleteFromDB
		'ExSummary:Delete the document from the database, using filename to fetch the record.
		Public Shared Sub DeleteFromDatabase(ByVal fileName As String)
			' Create the SQL command.
			Dim commandString As String = "DELETE * FROM Documents WHERE FileName='" & fileName & "'"
			Dim command As New OleDbCommand(commandString, mConnection)

			' Delete the record.
			command.ExecuteNonQuery()

		End Sub
		'ExEnd

		''' <summary>
		''' Connection to the database.
		''' </summary>
		Public Shared mConnection As OleDbConnection
	End Class
End Namespace