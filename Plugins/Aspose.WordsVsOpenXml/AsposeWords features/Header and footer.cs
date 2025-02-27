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
    public class ChangeOrReplaceHeaderAndFooter : TestUtil
    {
        [Test]
        public void CreateHeaderFooterAsposeWords()
        {
            //ExStart:CreateHeaderFooterAsposeWords
            //GistId:5cd48a22126e62b3a7a964491a234473
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Aspose.Words Header");
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write("Aspose.Words Footer");

            doc.Save(ArtifactsDir + "Create header footer - Aspose.Words.docx");
            //ExEnd:CreateHeaderFooterAsposeWords
        }
    }
}
