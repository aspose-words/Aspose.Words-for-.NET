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
    public class CreateTable : TestUtil
    {
        [Test]
        public void TableAsposeWords()
        {
            //ExStart:TableAsposeWords
            //GistId:657c3f65a58ea6dba182a07a830bcf4f
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

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

            doc.Save(ArtifactsDir + "Table - Aspose.Words.docx");
            //ExEnd:TableAsposeWords
        }
    }
}
