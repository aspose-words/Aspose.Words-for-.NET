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
    public class ChangeParagraphStyle : TestUtil
    {
        [Test]
        public void ParagraphCustomStyleAsposeWords()
        {
            //ExStart:ParagraphCustomStyleAsposeWords
            //GistId:bb3d63e124a55605dff971757e269bdc
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

            builder.MoveToDocumentEnd();

            builder.ParagraphFormat.Style = doc.Styles["MyStyle"];
            builder.Writeln("This string is formatted using the new style.");

            doc.Save(ArtifactsDir + "Paragraph custom style - Aspose.Words.docx");
            //ExEnd:ParagraphCustomStyleAsposeWords
        }
    }
}
