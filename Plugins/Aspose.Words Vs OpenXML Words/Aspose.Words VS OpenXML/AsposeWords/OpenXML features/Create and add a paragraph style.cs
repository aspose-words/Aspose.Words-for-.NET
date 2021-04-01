// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class CreateAndAddAParagraphStyle : TestUtil
    {
        [Test]
        public void CreateAndAddAParagraphStyleFeature()
        {
            using (WordprocessingDocument doc =
                WordprocessingDocument.Create(ArtifactsDir + "Create and add a paragraph style - OpenXML.docx",
                    WordprocessingDocumentType.Document))
            {
                // Get the Styles part for this document.
                // If the Styles part does not exist: add the styles part, and then add the style.
                StyleDefinitionsPart part =
                    doc.MainDocumentPart.StyleDefinitionsPart ?? AddStylesPartToPackage(doc);

                // Set up a variable to hold the style ID.
                string parastyleId = "OverdueAmountPara";

                // Create and add a paragraph style to the specified styles part 
                // with the specified style ID, style name, and aliases.
                AddParagraphStyle(part,
                    parastyleId,
                    "Overdue Amount Para",
                    "Late Due, Late Amount");

                // Add a paragraph with a run and some text.
                Paragraph p =
                    new Paragraph(
                        new Run(
                            new Text("This is some text in a run in a paragraph.")));

                // Add the paragraph as a child element of the w:body element.
                doc.MainDocumentPart.Document.Body.AppendChild(p);

                // If the paragraph has no ParagraphProperties object, then create one.
                if (!p.Elements<ParagraphProperties>().Any())
                {
                    p.PrependChild(new ParagraphProperties());
                }

                // Get a reference to the ParagraphProperties object.
                ParagraphProperties pPr = p.ParagraphProperties;

                // If a ParagraphStyleId object doesn't exist, then create one.
                if (pPr.ParagraphStyleId == null)
                    pPr.ParagraphStyleId = new ParagraphStyleId();

                pPr.ParagraphStyleId.Val = parastyleId;
            }
        }

        // Create a new paragraph style with the specified style ID, primary style name, and aliases.
        // Add it to the specified style definitions part.
        public static void AddParagraphStyle(StyleDefinitionsPart styleDefinitionsPart,
            string styleid, string stylename, string aliases = "")
        {
            // Access the root element of the styles part.
            Styles styles = styleDefinitionsPart.Styles;
            if (styles == null)
            {
                styleDefinitionsPart.Styles = new Styles();
                styleDefinitionsPart.Styles.Save();
            }

            // Create a new paragraph style element and specify some of the attributes.
            Style style = new Style
            {
                Type = StyleValues.Paragraph,
                StyleId = styleid,
                CustomStyle = true,
                Default = false
            };

            // Create and add the child elements (properties of the style).
            Aliases aliases1 = new Aliases { Val = aliases };
            AutoRedefine autoredefine1 = new AutoRedefine { Val = OnOffOnlyValues.Off };
            BasedOn basedon1 = new BasedOn { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle { Val = "OverdueAmountChar" };
            Locked locked1 = new Locked { Val = OnOffOnlyValues.Off };
            PrimaryStyle primarystyle1 = new PrimaryStyle { Val = OnOffOnlyValues.On };
            StyleHidden stylehidden1 = new StyleHidden { Val = OnOffOnlyValues.Off };
            SemiHidden semihidden1 = new SemiHidden { Val = OnOffOnlyValues.Off };
            StyleName styleName1 = new StyleName { Val = stylename };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle { Val = "Normal" };
            UIPriority uipriority1 = new UIPriority { Val = 1 };
            UnhideWhenUsed unhidewhenused1 = new UnhideWhenUsed { Val = OnOffOnlyValues.On };
            
            if (aliases != "")
                style.Append(aliases1);
            style.Append(autoredefine1);
            style.Append(basedon1);
            style.Append(linkedStyle1);
            style.Append(locked1);
            style.Append(primarystyle1);
            style.Append(stylehidden1);
            style.Append(semihidden1);
            style.Append(styleName1);
            style.Append(nextParagraphStyle1);
            style.Append(uipriority1);
            style.Append(unhidewhenused1);

            // Create the StyleRunProperties object and specify some of the run properties.
            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Bold bold1 = new Bold();
            Color color1 = new Color { ThemeColor = ThemeColorValues.Accent2 };
            RunFonts font1 = new RunFonts { Ascii = "Lucida Console" };
            Italic italic1 = new Italic();

            // Specify a 12 point size.
            FontSize fontSize1 = new FontSize { Val = "24" };
            styleRunProperties1.Append(bold1);
            styleRunProperties1.Append(color1);
            styleRunProperties1.Append(font1);
            styleRunProperties1.Append(fontSize1);
            styleRunProperties1.Append(italic1);

            // Add the run properties to the style.
            style.Append(styleRunProperties1);

            // Add the style to the styles part.
            styles.Append(style);
        }

        // Add a StylesDefinitionsPart to the document.
        public static StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument doc)
        {
            Styles root = new Styles();
            
            var part = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            root.Save(part);
            
            return part;
        }
    }
}
