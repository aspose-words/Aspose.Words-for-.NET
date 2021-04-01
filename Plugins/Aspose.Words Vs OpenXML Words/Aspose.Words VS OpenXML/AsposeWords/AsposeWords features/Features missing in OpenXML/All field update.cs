// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features.Features_missing_in_OpenXML
{
    [TestFixture]
    public class AllFieldUpdate : TestUtil
    {
        [Test]
        public void AllFieldUpdateFeature()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table of contents for the first page of the document.
            // Configure the table to pick up paragraphs with headings of levels 1 to 3.
            // Also, set its entries to be hyperlinks that will take us
            // to the location of the heading when left-clicked in Microsoft Word.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            builder.InsertBreak(BreakType.PageBreak);

            // Populate the table of contents by adding paragraphs with heading styles.
            // Each such heading with a level between 1 and 3 will create an entry in the table.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Heading 1.1");
            builder.Writeln("Heading 1.2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Heading 2");
            builder.Writeln("Heading 3");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Heading 3.1");

            // A table of contents is a field of a type that needs to be updated to show an up-to-date result.
            doc.UpdateFields();
            doc.Save(ArtifactsDir + "All field update - Aspose.Words.docx");
        }
    }
}
