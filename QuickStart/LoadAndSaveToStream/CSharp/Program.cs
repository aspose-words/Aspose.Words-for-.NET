//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Words;

namespace LoadAndSaveToStreamExample
{
    public class Program
    {
        public static void Main()
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // Open the stream. Read only access is enough for Aspose.Words to load a document.
            Stream stream = File.OpenRead(dataDir + "Document.doc");

            // Load the entire document into memory.
            Document doc = new Document(stream);

            // You can close the stream now, it is no longer needed because the document is in memory.
            stream.Close();

            // ... do something with the document

            // Convert the document to a different format and save to stream.
            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Rtf);

            // Rewind the stream position back to zero so it is ready for the next reader.
            dstStream.Position = 0;

            // Save the document from stream, to disk. Normally you would do something with the stream directly,
            // for example writing the data to a database.
            File.WriteAllBytes(dataDir + "Document Out.rtf", dstStream.ToArray());
        }
    }
}