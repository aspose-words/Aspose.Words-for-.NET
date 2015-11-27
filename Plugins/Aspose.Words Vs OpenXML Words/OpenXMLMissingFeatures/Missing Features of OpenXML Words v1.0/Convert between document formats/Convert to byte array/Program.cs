// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System.IO;
using Aspose.Words;

namespace ConvertToByteArray
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = @"Files\";
            // Load the document.
            Document doc = new Document(MyDir + "Converting Document.docx");

            // Create a new memory stream.
            MemoryStream outStream = new MemoryStream();
            // Save the document to stream.
            doc.Save(outStream, SaveFormat.Docx);

            // Convert the document to byte form.
            byte[] docBytes = outStream.ToArray();

            // The bytes are now ready to be stored/transmitted.

            // Now reverse the steps to load the bytes back into a document object.
            MemoryStream inStream = new MemoryStream(docBytes);

            // Load the stream into a new document object.
            Document loadDoc = new Document(inStream);
        }
    }
}
