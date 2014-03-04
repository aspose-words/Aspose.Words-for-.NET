//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Reflection;

using Aspose.Words;

namespace DocumentInDBExample
{
    /// <summary>
    /// This project is set to target the x86 platform because there is no 64-bit driver available 
    /// for the Access database used in this sample.
    /// </summary>
    public class Program
    {
        public static void Main()
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");
            string dbName = dataDir + "DocDB.mdb";
            const string fileName = "TestFile.doc";

            //ExStart
            //ExId:DocumentInDB_DatabaseHelpers
            //ExSummary:Shows how to setup a connection to a database and execute commands.
            // Create a connection to the database.
            mConnection = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dbName);
            
            // Open the database connection.
            mConnection.Open();
            //ExEnd

            // Open the document.
            Document doc = new Document(dataDir + fileName);

            //ExStart
            //ExId:DocumentInDB_Main
            //ExSummary:Stores the document to a database, then reads the same document back again, and finally deletes the record containing the document from the database.
            // Store the document to the database.
            StoreToDatabase(doc);
            // Read the document from the database and store the file to disk.
            Document dbDoc = ReadFromDatabase(fileName, dataDir);

            // Save the retrieved document to disk.
            string newFileName = Path.GetFileNameWithoutExtension(fileName) + " from DB" + Path.GetExtension(fileName);
            dbDoc.Save(dataDir + newFileName);

            // Delete the document from the database.
            DeleteFromDatabase(fileName);
            
            // Close the connection to the database.
            mConnection.Close();
            //ExEnd
        }

        /// <summary>
        /// Stores a document object to the specified database.
        /// </summary>
        /// <param name="dbName">The name of the database file.</param>
        /// <param name="doc">The source document.</param>
        /// 
        //ExStart
        //ExId:DocumentInDB_StoreToDB
        //ExSummary:Stores the document to the specified database.
        public static void StoreToDatabase(Document doc)
        {
            // Save the document to a MemoryStream object.
            MemoryStream stream = new MemoryStream();
            doc.Save(stream, SaveFormat.Doc);

            // Get the filename from the document.
            string fileName = Path.GetFileName(doc.OriginalFileName);

            // Create the SQL command.
            string commandString = "INSERT INTO Documents (FileName, FileContent) VALUES('" + fileName + "', @Doc)";
            OleDbCommand command = new OleDbCommand(commandString, mConnection);

            // Add the @Doc parameter.
            command.Parameters.AddWithValue("Doc", stream.ToArray());
            
            // Write the document to the database.
            command.ExecuteNonQuery();

        }
        //ExEnd

        /// <summary>
        /// Retreives a document from the specified database and saves it to disk.
        /// </summary>
        /// <param name="dbName">The name of the database file.</param>
        /// <param name="fileName">The name of the document file.</param>
        /// <param name="path">The path to the directory where to extract the document.</param>
        //ExStart
        //ExId:DocumentInDB_ReadFromDB
        //ExSummary:Retrieves and returns the document from the specified database using the filename as a key to fetch the document.
        public static Document ReadFromDatabase(string fileName, string path)
        {
            // Create the SQL command.
            string commandString = "SELECT * FROM Documents WHERE FileName='" + fileName + "'";
            OleDbCommand command = new OleDbCommand(commandString, mConnection);
           
            // Create the data adapter.
            OleDbDataAdapter adapter = new OleDbDataAdapter(command);

            // Fill the results from the database into a DataTable.
            DataTable dataTable = new DataTable();
            adapter.Fill(dataTable);

            // Check there was a matching record found from the database and throw an exception if no record was found.
            if (dataTable.Rows.Count == 0)
                throw new ArgumentException(string.Format("Could not find any record matching the document \"{0}\" in the database.", fileName));
            
            // The document is stored in byte form in the FileContent column.
            // Retrieve these bytes of the first matching record to a new buffer.
            byte[] buffer = (byte[])dataTable.Rows[0]["FileContent"];

            // Wrap the bytes from the buffer into a new MemoryStream object.
            MemoryStream newStream = new MemoryStream(buffer);

            // Read the document from the stream.
            Document doc = new Document(newStream);

            // Return the retrieved document.
            return doc;
            
        }
        //ExEnd

        /// <summary>
        /// Deletes the records containing the specified document name from the database.
        /// </summary>
        /// <param name="dbName">The name of the database file.</param>
        /// <param name="fileName">The name of the document file.</param>
        //ExStart
        //ExId:DocumentInDB_DeleteFromDB
        //ExSummary:Delete the document from the database, using filename to fetch the record.
        public static void DeleteFromDatabase(string fileName)
        {
            // Create the SQL command.
            string commandString = "DELETE * FROM Documents WHERE FileName='" + fileName + "'";
            OleDbCommand command = new OleDbCommand(commandString, mConnection);

            // Delete the record.
            command.ExecuteNonQuery();

        }
        //ExEnd

        /// <summary>
        /// Connection to the database.
        /// </summary>
        public static OleDbConnection mConnection;
    }
}