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
    public class FindAndReplaceText : TestUtil
    {
        [Test]
        public static void FindAndReplaceTextAsposeWords()
        {
            //ExStart:FindAndReplaceTextAsposeWords
            //GistDesc:Find and replace text in a document using C#
            Document doc = new Document(MyDir + "Search and replace text.docx");

            doc.Range.Replace("Hello World!", "Hi Everyone!");

            doc.Save(ArtifactsDir + "Find and replace text - Aspose.Words.docx");
            //ExEnd:FindAndReplaceTextAsposeWords
        }
    }
}
