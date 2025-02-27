// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class ChangeParagraphStyle : TestUtil
    {
        [Test]
        public void ParagraphCustomStyleOpenXml()
        {
            //ExStart:ParagraphCustomStyleOpenXml
            //GistId:bb3d63e124a55605dff971757e269bdc
            // Create a Wordprocessing document.
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(
                ArtifactsDir + "Paragraph custom style - OpenXML.docx",
                WordprocessingDocumentType.Document))
            {
                // Add a main document part.
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = new Body();

                // Create a custom paragraph style.
                StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
                Styles styles = new Styles();

                // Define a new paragraph style.
                Style paragraphStyle = new Style()
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "CustomStyle",
                    CustomStyle = true,
                    StyleName = new StyleName() { Val = "Custom Style" },
                    BasedOn = new BasedOn() { Val = "Normal" },
                };

                styles.Append(paragraphStyle);
                stylePart.Styles = styles;
                stylePart.Styles.Save();

                // Create a new paragraph with the custom style.
                Paragraph paragraph = new Paragraph()
                {
                    ParagraphProperties = new ParagraphProperties()
                    {
                        ParagraphStyleId = new ParagraphStyleId() { Val = "CustomStyle" },
                        Justification = new Justification() { Val = JustificationValues.Center },
                        SpacingBetweenLines = new SpacingBetweenLines() { After = "200" }
                    }
                };
                Run run = new Run();

                // Set bold text.
                RunProperties runProperties = new RunProperties();
                runProperties.Append(new Bold());
                run.Append(runProperties);
                run.Append(new Text("This is a bold paragraph with a custom style!"));
                paragraph.Append(run);
                body.Append(paragraph);

                mainPart.Document.Append(body);
                mainPart.Document.Save();
                //ExEnd:ParagraphCustomStyleOpenXml
            }
        }
    }
}
