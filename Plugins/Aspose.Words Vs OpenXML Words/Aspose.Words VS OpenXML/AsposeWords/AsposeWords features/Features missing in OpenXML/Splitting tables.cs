// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using Aspose.Words;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features.Features_missing_in_OpenXML
{
    [TestFixture]
    public class SplittingTables : TestUtil
    {
        [Test]
        public static void SplittingTablesFeature()
        {
            Document doc = new Document(MyDir + "Tables.docx");

            // Get the first table in the document.
            Table firstTable = (Table)doc.GetChild(NodeType.Table, 0, true);

            // We will split the table at the third row (inclusive).
            Row row = firstTable.Rows[2];

            // Create a new container for the split table.
            Table table = (Table)firstTable.Clone(false);

            // Insert the container after the original.
            firstTable.ParentNode.InsertAfter(table, firstTable);

            // Add a buffer paragraph to ensure the tables stay apart.
            firstTable.ParentNode.InsertAfter(new Paragraph(doc), firstTable);

            Row currentRow;

            do
            {
                currentRow = firstTable.LastRow;
                table.PrependChild(currentRow);
            }
            while (currentRow != row);

            doc.Save(ArtifactsDir + "Splitting tables - Aspose.Words.docx");
        }
    }
}
