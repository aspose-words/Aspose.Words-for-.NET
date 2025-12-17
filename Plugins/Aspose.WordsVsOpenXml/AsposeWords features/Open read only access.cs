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
    class OpenReadOnlyAccess : TestUtil
    {
        [Test]
        public void OpenReadOnlyAsposeWords()
        {
            //ExStart:OpenReadOnlyAsposeWords
            //GistId:702c287894827f3d4ddd2ca4b170ed45
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Open document as read-only");

            // Enter a password that's up to 15 characters long.
            doc.WriteProtection.SetPassword("MyPassword");
            // Make the document as read-only.
            doc.WriteProtection.ReadOnlyRecommended = true;

            // Apply write protection as read-only.
            doc.Protect(ProtectionType.ReadOnly);
            doc.Save(ArtifactsDir + "ReadOnly protection - Aspose.Words.docx");
            //ExEnd:OpenReadOnlyAsposeWords
        }
    }
}
