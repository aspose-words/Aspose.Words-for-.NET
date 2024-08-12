﻿using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Complex_examples_and_helpers
{
    public class WorkingWithDocumentInDatabase : DocsExamplesBase
    {
#if NET48 || JAVA
        [Test, Category("IgnoreOnJenkins")]
        public void LoadAndSaveDocToDatabase()
        {
            Document doc = new Document(MyDir + "Document.docx");
            //ExStart:OpenDatabaseConnection
            //GistId:f8a622f8bc1cf3c2fa8a7a9be359faa2
            string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + DatabaseDir + "Northwind.accdb";
            
            OleDbConnection connection = new OleDbConnection(connString);
            connection.Open();
            //ExEnd:OpenDatabaseConnection

            //ExStart:OpenRetrieveAndDelete
            //GistId:f8a622f8bc1cf3c2fa8a7a9be359faa2
            StoreToDatabase(doc, connection);
            
            Document dbDoc = ReadFromDatabase("Document.docx", connection);
            dbDoc.Save(ArtifactsDir + "WorkingWithDocumentInDatabase.LoadAndSaveDocToDatabase.docx");

            DeleteFromDatabase("Document.docx", connection);

            connection.Close();
            //ExEnd:OpenRetrieveAndDelete
        }

        //ExStart:StoreToDatabase
        //GistId:f8a622f8bc1cf3c2fa8a7a9be359faa2
        public void StoreToDatabase(Document doc, OleDbConnection connection)
        {
            MemoryStream stream = new MemoryStream();
            doc.Save(stream, SaveFormat.Docx);

            string fileName = Path.GetFileName(doc.OriginalFileName);
            string commandString = "INSERT INTO Documents (Name, Data) VALUES('" + fileName + "', @Doc)";
            
            OleDbCommand command = new OleDbCommand(commandString, connection);
            command.Parameters.AddWithValue("Doc", stream.ToArray());
            command.ExecuteNonQuery();
        }
        //ExEnd:StoreToDatabase

        //ExStart:ReadFromDatabase
        //GistId:f8a622f8bc1cf3c2fa8a7a9be359faa2
        public Document ReadFromDatabase(string fileName, OleDbConnection connection)
        {
            string commandString = "SELECT * FROM Documents WHERE Name='" + fileName + "'";
            
            OleDbCommand command = new OleDbCommand(commandString, connection);
            OleDbDataAdapter adapter = new OleDbDataAdapter(command);

            DataTable dataTable = new DataTable();
            adapter.Fill(dataTable);

            if (dataTable.Rows.Count == 0)
                throw new ArgumentException(
                    $"Could not find any record matching the document \"{fileName}\" in the database.");

            // The document is stored in byte form in the FileContent column.
            // Retrieve these bytes of the first matching record to a new buffer.
            byte[] buffer = (byte[]) dataTable.Rows[0]["Data"];

            MemoryStream newStream = new MemoryStream(buffer);

            Document doc = new Document(newStream);

            return doc;
        }
        //ExEnd:ReadFromDatabase

        //ExStart:DeleteFromDatabase
        //GistId:f8a622f8bc1cf3c2fa8a7a9be359faa2
        public void DeleteFromDatabase(string fileName, OleDbConnection connection)
        {
            string commandString = "DELETE * FROM Documents WHERE Name='" + fileName + "'";
            
            OleDbCommand command = new OleDbCommand(commandString, connection);
            command.ExecuteNonQuery();
        }
        //ExEnd:DeleteFromDatabase
#endif
    }
}