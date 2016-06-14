
using System.IO;

using Aspose.Words;
using System;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class LoadAndSaveToStream
    {
        public static void Run()
        {
            //ExStart:LoadAndSaveToStream 
            //ExStart:OpeningFromStream 
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_QuickStart();
            string fileName = "Document.doc";

            // Open the stream. Read only access is enough for Aspose.Words to load a document.
            Stream stream = File.OpenRead(dataDir + fileName);

            // Load the entire document into memory.
            Document doc = new Document(stream);

            // You can close the stream now, it is no longer needed because the document is in memory.
            stream.Close();
            //ExEnd:OpeningFromStream 

            // ... do something with the document

            // Convert the document to a different format and save to stream.
            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Rtf);

            // Rewind the stream position back to zero so it is ready for the next reader.
            dstStream.Position = 0;
            //ExEnd:LoadAndSaveToStream 
            // Save the document from stream, to disk. Normally you would do something with the stream directly,
            // for example writing the data to a database.
            dataDir = dataDir + "Document_out_.rtf";
            File.WriteAllBytes(dataDir, dstStream.ToArray());

            Console.WriteLine("\nStream of document saved successfully.\nFile saved at " + dataDir);
        }
    }
}
