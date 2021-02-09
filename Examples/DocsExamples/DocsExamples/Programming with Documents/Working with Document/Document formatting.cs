using System;
using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Working_with_Document
{
    internal class DocumentFormatting : DocsExamplesBase
    {
        [Test]
        public void SpaceBetweenAsianAndLatinText()
        {
            //ExStart:SpaceBetweenAsianAndLatinText
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            ParagraphFormat paragraphFormat = builder.ParagraphFormat;
            paragraphFormat.AddSpaceBetweenFarEastAndAlpha = true;
            paragraphFormat.AddSpaceBetweenFarEastAndDigit = true;

            builder.Writeln("Automatically adjust space between Asian and Latin text");
            builder.Writeln("Automatically adjust space between Asian text and numbers");

            doc.Save(ArtifactsDir + "DocumentFormatting.SpaceBetweenAsianAndLatinText.docx");
            //ExEnd:SpaceBetweenAsianAndLatinText
        }

        [Test]
        public void AsianTypographyLineBreakGroup()
        {
            //ExStart:AsianTypographyLineBreakGroup
            Document doc = new Document(MyDir + "Asian typography.docx");

            ParagraphFormat format = doc.FirstSection.Body.Paragraphs[0].ParagraphFormat;
            format.FarEastLineBreakControl = false;
            format.WordWrap = true;
            format.HangingPunctuation = false;

            doc.Save(ArtifactsDir + "DocumentFormatting.AsianTypographyLineBreakGroup.docx");
            //ExEnd:AsianTypographyLineBreakGroup
        }

        [Test]
        public void ParagraphFormatting()
        {
            //ExStart:ParagraphFormatting
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            ParagraphFormat paragraphFormat = builder.ParagraphFormat;
            paragraphFormat.Alignment = ParagraphAlignment.Center;
            paragraphFormat.LeftIndent = 50;
            paragraphFormat.RightIndent = 50;
            paragraphFormat.SpaceAfter = 25;

            builder.Writeln(
                "I'm a very nice formatted paragraph. I'm intended to demonstrate how the left and right indents affect word wrapping.");
            builder.Writeln(
                "I'm another nice formatted paragraph. I'm intended to demonstrate how the space after paragraph looks like.");

            doc.Save(ArtifactsDir + "DocumentFormatting.ParagraphFormatting.docx");
            //ExEnd:ParagraphFormatting
        }

        [Test]
        public void MultilevelListFormatting()
        {
            //ExStart:MultilevelListFormatting
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");

            builder.ListFormat.ListIndent();
            builder.Writeln("Item 2.1");
            builder.Writeln("Item 2.2");
            
            builder.ListFormat.ListIndent();
            builder.Writeln("Item 2.2.1");
            builder.Writeln("Item 2.2.2");

            builder.ListFormat.ListOutdent();
            builder.Writeln("Item 2.3");

            builder.ListFormat.ListOutdent();
            builder.Writeln("Item 3");

            builder.ListFormat.RemoveNumbers();
            
            doc.Save(ArtifactsDir + "DocumentFormatting.MultilevelListFormatting.docx");
            //ExEnd:MultilevelListFormatting
        }

        [Test]
        public void ApplyParagraphStyle()
        {
            //ExStart:ApplyParagraphStyle
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
            builder.Write("Hello");
            
            doc.Save(ArtifactsDir + "DocumentFormatting.ApplyParagraphStyle.docx");
            //ExEnd:ApplyParagraphStyle
        }

        [Test]
        public void ApplyBordersAndShadingToParagraph()
        {
            //ExStart:ApplyBordersAndShadingToParagraph
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            BorderCollection borders = builder.ParagraphFormat.Borders;
            borders.DistanceFromText = 20;
            borders[BorderType.Left].LineStyle = LineStyle.Double;
            borders[BorderType.Right].LineStyle = LineStyle.Double;
            borders[BorderType.Top].LineStyle = LineStyle.Double;
            borders[BorderType.Bottom].LineStyle = LineStyle.Double;

            Shading shading = builder.ParagraphFormat.Shading;
            shading.Texture = TextureIndex.TextureDiagonalCross;
            shading.BackgroundPatternColor = System.Drawing.Color.LightCoral;
            shading.ForegroundPatternColor = System.Drawing.Color.LightSalmon;

            builder.Write("I'm a formatted paragraph with double border and nice shading.");
            
            doc.Save(ArtifactsDir + "DocumentFormatting.ApplyBordersAndShadingToParagraph.doc");
            //ExEnd:ApplyBordersAndShadingToParagraph
        }
        
        [Test]
        public void ChangeAsianParagraphSpacingAndIndents()
        {
            //ExStart:ChangeAsianParagraphSpacingAndIndents
            Document doc = new Document(MyDir + "Asian typography.docx");

            ParagraphFormat format = doc.FirstSection.Body.FirstParagraph.ParagraphFormat;
            format.CharacterUnitLeftIndent = 10;       // ParagraphFormat.LeftIndent will be updated
            format.CharacterUnitRightIndent = 10;      // ParagraphFormat.RightIndent will be updated
            format.CharacterUnitFirstLineIndent = 20;  // ParagraphFormat.FirstLineIndent will be updated
            format.LineUnitBefore = 5;                 // ParagraphFormat.SpaceBefore will be updated
            format.LineUnitAfter = 10;                 // ParagraphFormat.SpaceAfter will be updated

            doc.Save(ArtifactsDir + "DocumentFormatting.ChangeAsianParagraphSpacingAndIndents.doc");
            //ExEnd:ChangeAsianParagraphSpacingAndIndents
        }

        [Test]
        public void SnapToGrid()
        {
            //ExStart:SetSnapToGrid
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Optimize the layout when typing in Asian characters.
            Paragraph par = doc.FirstSection.Body.FirstParagraph;
            par.ParagraphFormat.SnapToGrid = true;

            builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod " +
                            "tempor incididunt ut labore et dolore magna aliqua.");
            
            par.Runs[0].Font.SnapToGrid = true;

            doc.Save(ArtifactsDir + "Paragraph.SnapToGrid.docx");
            //ExEnd:SetSnapToGrid
        }

        [Test]
        public void GetParagraphStyleSeparator()
        {
            //ExStart:GetParagraphStyleSeparator
            Document doc = new Document(MyDir + "Document.docx");

            foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (paragraph.BreakIsStyleSeparator)
                {
                    Console.WriteLine("Separator Found!");
                }
            }
            //ExEnd:GetParagraphStyleSeparator
        }
    }
}