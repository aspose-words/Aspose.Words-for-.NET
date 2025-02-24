// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    class ChangeTextInATable : TestUtil
    {
        [Test]
        public void ReplaceText()
        {
            Document doc = new Document(MyDir + "Change text in a table.docx");

            // Get the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            // Replace any instances of our string in the last cell of the table only.
            FindReplaceOptions options = new FindReplaceOptions
            {
                MatchCase = true, 
                FindWholeWordsOnly = true
            };
            table.Rows[1].Cells[2].Range.Replace("Mr", "test", options);

            doc.Save(ArtifactsDir + "Replace text - Aspose.Words.docx");
        }
    }
}
