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
    public class ConvertFromDocmToDocx : TestUtil
    {
        [Test]
        public void DocmToDocxAsposeWords()
        {
            //ExStart:DocmToDocxAsposeWords
            //GistDesc:Convert Docm to Docx using C#
            Document doc = new Document(MyDir + "Docm to Docx.docm");
            doc.Save(ArtifactsDir + "Docm to Docx - Aspose.Words.docx");
            //ExEnd:DocmToDocxAsposeWords
        }
    }
}
