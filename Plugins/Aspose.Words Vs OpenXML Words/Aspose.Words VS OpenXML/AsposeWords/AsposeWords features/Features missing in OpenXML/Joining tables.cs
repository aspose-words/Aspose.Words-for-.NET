// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using Aspose.Words;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features.Features_missing_in_OpenXML
{
    [TestFixture]
    public class JoiningTables : TestUtil
    {
        [Test]
        public static void JoiningTablesFeatures()
        {
            Document doc = new Document(MyDir + "Tables.docx");

            // Get the first and second table in the document.
            // The rows from the second table will be appended to the end of the first table.
            Table firstTable = (Table)doc.GetChild(NodeType.Table, 0, true);
            Table secondTable = (Table)doc.GetChild(NodeType.Table, 1, true);

            // Append all rows from the current table to the next.
            // Due to the design of tables even tables with different cell count and widths can be joined into one table.
            while (secondTable.HasChildNodes)
                firstTable.Rows.Add(secondTable.FirstRow);

            // Remove the empty table container.
            secondTable.Remove();

            doc.Save(ArtifactsDir + "Joining tables - Aspose.Words.docx");
        }
    }
}
