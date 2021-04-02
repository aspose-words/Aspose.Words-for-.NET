// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System.IO;
using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features.Features_missing_in_OpenXML
{
    [TestFixture]
    public class ConvertToByteArray : TestUtil
    {
        [Test]
        public static void ConvertToByteArrayFeature()
        {
            Document doc = new Document(MyDir + "Document.docx");

            MemoryStream outStream = new MemoryStream();
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
