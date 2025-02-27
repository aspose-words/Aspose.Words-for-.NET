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
    public class RemoveHeaderFooter : TestUtil
    {
        [Test]
        public void RemoveHeaderFooterAsposeWords()
        {
            //ExStart:RemoveHeaderFooterAsposeWords
            //GistDesc:Remove headers and footers using C#
            Document doc = new Document(MyDir + "Document.docx");

            foreach (HeaderFooter headerFooter in doc.GetChildNodes(NodeType.HeaderFooter, true))
                headerFooter.Remove();

            doc.Save(ArtifactsDir + "Remove header and footer - Aspose.Words.docx");
            //ExEnd:RemoveHeaderFooterAsposeWords
        }
    }
}
