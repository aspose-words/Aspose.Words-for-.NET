// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class AddTable : TestUtil
    {
        [Test]
        public void AddTableFeature()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a 2x2 table.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Row 1, Cell 1 Content.");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2 Content.");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Row 2, Cell 1 Content");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2 Content.");
            builder.EndRow();
            builder.EndTable();

            doc.Save(ArtifactsDir + "Add table - Aspose.Words.docx");
        }
    }
}
