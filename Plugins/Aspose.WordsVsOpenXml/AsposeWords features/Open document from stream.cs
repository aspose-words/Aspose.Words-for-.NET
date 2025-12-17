// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class OpenDocumentFromStream : TestUtil
    {
        [Test]
        public void AddTextStreamAsposeWords()
        {
            //ExStart:AddTextStreamAsposeWords
            //GistId:a230affc64d19e458a3a6a5452903946
            using (Stream stream = File.Open(MyDir + "Document.docx", FileMode.Open))
            {
                Document doc = new Document(stream);
                DocumentBuilder builder = new DocumentBuilder(doc);

                builder.Writeln();
                builder.Write("This is the text added to the end of the document.");
                
                doc.Save(ArtifactsDir + "Add text stream - Aspose.Words.docx");
            }
            //ExEnd:AddTextStreamAsposeWords
        }
    }
}
