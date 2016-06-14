using Aspose.Words;
using Aspose.Words.MailMerging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;

namespace Aspose.Words.Examples.CSharp.Mail_Merge
{
    class MailMergeImageFromBlob
    {
        public static void Run()
        {
            //ExStart:MailMergeImageFromBlob            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_MailMergeAndReporting(); 
            Document doc = new Document(dataDir + "MailMerge.MergeImage.doc");

            // Set up the event handler for image fields.
            doc.MailMerge.FieldMergingCallback = new HandleMergeImageFieldFromBlob();

            // Open a database connection.
            string connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + RunExamples.GetDataDir_Database()+"Northwind.mdb";
            OleDbConnection conn = new OleDbConnection(connString);
            conn.Open();

            // Open the data reader. It needs to be in the normal mode that reads all record at once.
            OleDbCommand cmd = new OleDbCommand("SELECT * FROM Employees", conn);
            IDataReader dataReader = cmd.ExecuteReader();

            // Perform mail merge.
            doc.MailMerge.ExecuteWithRegions(dataReader, "Employees");

            // Close the database.
            conn.Close();
            dataDir = dataDir + "MailMerge.MergeImage_out_.doc";
            doc.Save(dataDir);
            //ExEnd:MailMergeImageFromBlob
            Console.WriteLine("\nMail merge image from blob performed successfully.\nFile saved at " + dataDir);
        }
        //ExStart:HandleMergeImageFieldFromBlob 
        public class HandleMergeImageFieldFromBlob : IFieldMergingCallback
        {
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
            {
                // Do nothing.
            }

            /// <summary>
            /// This is called when mail merge engine encounters Image:XXX merge field in the document.
            /// You have a chance to return an Image object, file name or a stream that contains the image.
            /// </summary>
            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs e)
            {
                // The field value is a byte array, just cast it and create a stream on it.
                MemoryStream imageStream = new MemoryStream((byte[])e.FieldValue);
                // Now the mail merge engine will retrieve the image from the stream.
                e.ImageStream = imageStream;
            }
        }
        //ExEnd:HandleMergeImageFieldFromBlob
    }
}
