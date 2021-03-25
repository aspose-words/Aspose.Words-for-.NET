using Aspose.Words;
using System.IO;

namespace Convert_a_Document_to_a_Byte_Array
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"..\..\..\..\Sample Files\";

            // Load a document from the local file system.
            Document doc = new Document(filePath + "MyDocument.docx");

            byte[] docBytes;

            // Create a new memory stream.
            using (MemoryStream outStream = new MemoryStream())
            {
                // Save the document to stream.
                doc.Save(outStream, SaveFormat.Docx);

                // Convert the document to byte form.
                docBytes = outStream.ToArray();
            }

            // The bytes are now ready to be stored/transmitted.
            // Now reverse the steps to load the bytes back into a document object.
            MemoryStream inStream = new MemoryStream(docBytes);

            // Load the stream into a new document object.
            Document loadDoc = new Document(inStream);
        }
    }
}
