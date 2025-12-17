// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class CreateADocument : TestUtil
    {
        [Test]
        public void CreateNewDocumentAsposeWords()
        {
            //ExStart:CreateNewDocumentAsposeWords
            //GistId:e75459ad5b9ea7ac4cbea10ab631a491
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello World!");

            doc.Save(ArtifactsDir + "Create new document - Aspose.Words.docx");
            //ExEnd:CreateNewDocumentAsposeWords
        }
    }
}
