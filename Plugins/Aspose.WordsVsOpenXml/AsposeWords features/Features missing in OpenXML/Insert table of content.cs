// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features.Features_missing_in_OpenXML
{
    [TestFixture]
    public class InsertTableOfContent
    {
        [Test]
        public static void InsertTableOfContentFeature()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table of contents at the beginning of the document.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

            // The newly inserted table of contents will be initially empty.
            // It needs to be populated by updating the fields in the document.
            doc.UpdateFields();
        }
    }
}
