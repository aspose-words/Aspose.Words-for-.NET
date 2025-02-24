// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Loading;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    class OpenReadOnlyAccess : TestUtil
    {
        [Test]
        public void OpenReadOnly()
        {
            Document doc = new Document(MyDir + "Open ReadOnly access.docx", new LoadOptions("1234"));
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln();
            builder.Write("This is the text added to the end of the document.");
            
            doc.Save(ArtifactsDir + "Open encrypted - Aspose.Words.docx");
        }
    }
}
