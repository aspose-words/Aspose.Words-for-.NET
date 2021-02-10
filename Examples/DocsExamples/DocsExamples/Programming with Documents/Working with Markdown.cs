using System;
using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents
{
    internal class WorkingWithMarkdown : DocsExamplesBase
    {
        [Test]
        public void CreateMarkdownDocument()
        {
            //ExStart:CreateMarkdownDocument
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Specify the "Heading 1" style for the paragraph.
            builder.ParagraphFormat.StyleName = "Heading 1";
            builder.Writeln("Heading 1");

            // Reset styles from the previous paragraph to not combine styles between paragraphs.
            builder.ParagraphFormat.StyleName = "Normal";

            // Insert horizontal rule.
            builder.InsertHorizontalRule();

            // Specify the ordered list.
            builder.InsertParagraph();
            builder.ListFormat.ApplyNumberDefault();

            // Specify the Italic emphasis for the text.
            builder.Font.Italic = true;
            builder.Writeln("Italic Text");
            builder.Font.Italic = false;

            // Specify the Bold emphasis for the text.
            builder.Font.Bold = true;
            builder.Writeln("Bold Text");
            builder.Font.Bold = false;

            // Specify the StrikeThrough emphasis for the text.
            builder.Font.StrikeThrough = true;
            builder.Writeln("StrikeThrough Text");
            builder.Font.StrikeThrough = false;

            // Stop paragraphs numbering.
            builder.ListFormat.RemoveNumbers();

            // Specify the "Quote" style for the paragraph.
            builder.ParagraphFormat.StyleName = "Quote";
            builder.Writeln("A Quote block");

            // Specify nesting Quote.
            Style nestedQuote = doc.Styles.Add(StyleType.Paragraph, "Quote1");
            nestedQuote.BaseStyleName = "Quote";
            builder.ParagraphFormat.StyleName = "Quote1";
            builder.Writeln("A nested Quote block");

            // Reset paragraph style to Normal to stop Quote blocks. 
            builder.ParagraphFormat.StyleName = "Normal";

            // Specify a Hyperlink for the desired text.
            builder.Font.Bold = true;
            // Note, the text of hyperlink can be emphasized.
            builder.InsertHyperlink("Aspose", "https://www.aspose.com", false);
            builder.Font.Bold = false;

            // Insert a simple table.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell1");
            builder.InsertCell();
            builder.Write("Cell2");
            builder.EndTable();

            // Save your document as a Markdown file.
            doc.Save(ArtifactsDir + "WorkingWithMarkdown.CreateMarkdownDocument.md");
            //ExEnd:CreateMarkdownDocument
        }

        [Test]
        public void ReadMarkdownDocument()
        {
            //ExStart:ReadMarkdownDocument
            Document doc = new Document(MyDir + "Quotes.md");

            // Let's remove Heading formatting from a Quote in the very last paragraph.
            Paragraph paragraph = doc.FirstSection.Body.LastParagraph;
            paragraph.ParagraphFormat.Style = doc.Styles["Quote"];

            doc.Save(ArtifactsDir + "WorkingWithMarkdown.ReadMarkdownDocument.md");
            //ExEnd:ReadMarkdownDocument
        }

        [Test]
        public void Emphases()
        {
            //ExStart:Emphases
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Markdown treats asterisks (*) and underscores (_) as indicators of emphasis.");
            builder.Write("You can write ");

            builder.Font.Bold = true;
            builder.Write("bold");

            builder.Font.Bold = false;
            builder.Write(" or ");

            builder.Font.Italic = true;
            builder.Write("italic");

            builder.Font.Italic = false;
            builder.Writeln(" text. ");

            builder.Write("You can also write ");
            builder.Font.Bold = true;

            builder.Font.Italic = true;
            builder.Write("BoldItalic");

            builder.Font.Bold = false;
            builder.Font.Italic = false;
            builder.Write("text.");

            builder.Document.Save(ArtifactsDir + "WorkingWithMarkdown.Emphases.md");
            //ExEnd:Emphases
        }

        [Test]
        public void Headings()
        {
            //ExStart:Headings
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // By default Heading styles in Word may have bold and italic formatting.
            // If we do not want the text to be emphasized, set these properties explicitly to false.
            builder.Font.Bold = false;
            builder.Font.Italic = false;

            builder.Writeln("The following produces headings:");
            builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
            builder.Writeln("Heading1");
            builder.ParagraphFormat.Style = doc.Styles["Heading 2"];
            builder.Writeln("Heading2");
            builder.ParagraphFormat.Style = doc.Styles["Heading 3"];
            builder.Writeln("Heading3");
            builder.ParagraphFormat.Style = doc.Styles["Heading 4"];
            builder.Writeln("Heading4");
            builder.ParagraphFormat.Style = doc.Styles["Heading 5"];
            builder.Writeln("Heading5");
            builder.ParagraphFormat.Style = doc.Styles["Heading 6"];
            builder.Writeln("Heading6");

            // Note that the emphases are also allowed inside Headings.
            builder.Font.Bold = true;
            builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
            builder.Writeln("Bold Heading1");

            doc.Save(ArtifactsDir + "WorkingWithMarkdown.Headings.md");
            //ExEnd:Headings
        }

        [Test]
        public void BlockQuotes()
        {
            //ExStart:BlockQuotes
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("We support blockquotes in Markdown:");
            
            builder.ParagraphFormat.Style = doc.Styles["Quote"];
            builder.Writeln("Lorem");
            builder.Writeln("ipsum");
            
            builder.ParagraphFormat.Style = doc.Styles["Normal"];
            builder.Writeln("The quotes can be of any level and can be nested:");
            
            Style quoteLevel3 = doc.Styles.Add(StyleType.Paragraph, "Quote2");
            builder.ParagraphFormat.Style = quoteLevel3;
            builder.Writeln("Quote level 3");
            
            Style quoteLevel4 = doc.Styles.Add(StyleType.Paragraph, "Quote3");
            builder.ParagraphFormat.Style = quoteLevel4;
            builder.Writeln("Nested quote level 4");
            
            builder.ParagraphFormat.Style = doc.Styles["Quote"];
            builder.Writeln();
            builder.Writeln("Back to first level");
            
            Style quoteLevel1WithHeading = doc.Styles.Add(StyleType.Paragraph, "Quote Heading 3");
            builder.ParagraphFormat.Style = quoteLevel1WithHeading;
            builder.Write("Headings are allowed inside Quotes");

            doc.Save(ArtifactsDir + "WorkingWithMarkdown.BlockQuotes.md");
            //ExEnd:BlockQuotes
        }

        [Test]
        public void HorizontalRule()
        {
            //ExStart:HorizontalRule
            DocumentBuilder builder = new DocumentBuilder(new Document());

            builder.Writeln("We support Horizontal rules (Thematic breaks) in Markdown:");
            builder.InsertHorizontalRule();

            builder.Document.Save(ArtifactsDir + "WorkingWithMarkdown.HorizontalRuleExample.md");
            //ExEnd:HorizontalRule
        }

        [Test]
        public void UseWarningSource()
        {
            //ExStart:UseWarningSourceMarkdown
            Document doc = new Document(MyDir + "Emphases markdown warning.docx");

            WarningInfoCollection warnings = new WarningInfoCollection();
            doc.WarningCallback = warnings;

            doc.Save(ArtifactsDir + "WorkingWithMarkdown.UseWarningSource.md");

            foreach (WarningInfo warningInfo in warnings)
            {
                if (warningInfo.Source == WarningSource.Markdown)
                    Console.WriteLine(warningInfo.Description);
            }
            //ExEnd:UseWarningSourceMarkdown
        }
    }
}