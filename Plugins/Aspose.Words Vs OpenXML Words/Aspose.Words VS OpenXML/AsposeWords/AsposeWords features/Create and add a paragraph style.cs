// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class CreateAndAddAParagraphStyle : TestUtil
    {
        [Test]
        public void CreateAndAddAParagraphStyleFeature()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            doc.Styles.Add(StyleType.Paragraph, "MyStyle");
            Font font = builder.Font;
            font.Bold = true;
            font.Color = System.Drawing.Color.Blue;
            font.Italic = true;
            font.Name = "Arial";
            font.Size = 24;
            font.Spacing = 5;
            font.Underline = Underline.Double;

            builder.ParagraphFormat.Style = doc.Styles["MyStyle"];

            builder.MoveToDocumentEnd();
            builder.Writeln("This string is formatted using the new style.");

            doc.Save(ArtifactsDir + "Create and add a paragraph style - Aspose.Words.docx");
        }
    }
}
