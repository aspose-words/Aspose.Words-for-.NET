using System;
using System.Drawing;
using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Contents_Management
{
    internal class WorkingWithStylesAndThemes : DocsExamplesBase
    {
        [Test]
        public void AccessStyles()
        {
            //ExStart:AccessStyles
            //GistId:a73b495f610523670f0847331ef4d6fc
            Document doc = new Document();

            string styleName = "";

            // Get styles collection from the document.
            StyleCollection styles = doc.Styles;
            foreach (Style style in styles)
            {
                if (styleName == "")
                {
                    styleName = style.Name;
                    Console.WriteLine(styleName);
                }
                else
                {
                    styleName = styleName + ", " + style.Name;
                    Console.WriteLine(styleName);
                }
            }
            //ExEnd:AccessStyles
        }

        [Test]
        public void CopyStyles()
        {
            //ExStart:CopyStyles
            //GistId:a73b495f610523670f0847331ef4d6fc
            Document doc = new Document();
            Document target = new Document(MyDir + "Rendering.docx");

            target.CopyStylesFromTemplate(doc);

            doc.Save(ArtifactsDir + "WorkingWithStylesAndThemes.CopyStyles.docx");
            //ExEnd:CopyStyles
        }

        [Test]
        public void GetThemeProperties()
        {
            //ExStart:GetThemeProperties
            //GistId:a73b495f610523670f0847331ef4d6fc
            Document doc = new Document();

            Aspose.Words.Themes.Theme theme = doc.Theme;

            Console.WriteLine(theme.MajorFonts.Latin);
            Console.WriteLine(theme.MinorFonts.EastAsian);
            Console.WriteLine(theme.Colors.Accent1);
            //ExEnd:GetThemeProperties
        }

        [Test]
        public void SetThemeProperties()
        {
            //ExStart:SetThemeProperties
            //GistId:a73b495f610523670f0847331ef4d6fc
            Document doc = new Document();

            Aspose.Words.Themes.Theme theme = doc.Theme;
            theme.MinorFonts.Latin = "Times New Roman";
            theme.Colors.Hyperlink = Color.Gold;
            //ExEnd:SetThemeProperties
        }

        [Test]
        public void InsertStyleSeparator()
        {
            //ExStart:InsertStyleSeparator
            //GistId:4b5526c3c0d9cad73e05fb4b18d2c3d2
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Style paraStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MyParaStyle");
            paraStyle.Font.Bold = false;
            paraStyle.Font.Size = 8;
            paraStyle.Font.Name = "Arial";

            // Append text with "Heading 1" style.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Write("Heading 1");
            builder.InsertStyleSeparator();

            // Append text with another style.
            builder.ParagraphFormat.StyleName = paraStyle.Name;
            builder.Write("This is text with some other formatting ");

            doc.Save(ArtifactsDir + "WorkingWithStylesAndThemes.InsertStyleSeparator.docx");
            //ExEnd:InsertStyleSeparator
        }

        [Test]
        public void CopyStyleDifferentDocument()
        {
            //ExStart:CopyStyleDifferentDocument
            //GistId:93b92a7e6f2f4bbfd9177dd7fcecbd8c
            Document srcDoc = new Document();

            // Create a custom style for the source document.
            Style srcStyle = srcDoc.Styles.Add(StyleType.Paragraph, "MyStyle");
            srcStyle.Font.Color = Color.Red;

            // Import the source document's custom style into the destination document.
            Document dstDoc = new Document();
            Style newStyle = dstDoc.Styles.AddCopy(srcStyle);

            // The imported style has an appearance identical to its source style.
            Assert.That(newStyle.Name, Is.EqualTo("MyStyle"));
            Assert.That(newStyle.Font.Color.ToArgb(), Is.EqualTo(Color.Red.ToArgb()));
            //ExEnd:CopyStyleDifferentDocument
        }
    }
}