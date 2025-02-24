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
        public void DocmToDocxConversion()
        {
            Document doc = new Document(MyDir + "Docm to Docx conversion.docm");
            doc.Save(ArtifactsDir + "Docm to Docx conversion - Aspose.Words.docx");
        }
    }
}
