// Copyright (c) Aspose 2002-2021. All Rights Reserved.

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
        public void ChangeTextInATableFeature()
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

            doc.Save(ArtifactsDir + "Change text in a table - Aspose.Words.docx");
        }
    }
}
