﻿using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents
{
    internal class WorkingWithMarkdown : DocsExamplesBase
    {
        [Test]
        public void BoldText()
        {
            //ExStart:BoldText
            //GistId:0697355b7f872839932388d269ed6a63
            // Use a document builder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder();

            // Make the text Bold.
            builder.Font.Bold = true;
            builder.Writeln("This text will be Bold");
            //ExEnd:BoldText
        }

        [Test]
        public void ItalicText()
        {
            //ExStart:ItalicText
            //GistId:0697355b7f872839932388d269ed6a63
            // Use a document builder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder();

            // Make the text Italic.
            builder.Font.Italic = true;
            builder.Writeln("This text will be Italic");
            //ExEnd:ItalicText
        }

        [Test]
        public void Strikethrough()
        {
            //ExStart:Strikethrough
            //GistId:0697355b7f872839932388d269ed6a63
            // Use a document builder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder();

            // Make the text Strikethrough.
            builder.Font.StrikeThrough = true;
            builder.Writeln("This text will be StrikeThrough");
            //ExEnd:Strikethrough
        }

        [Test]
        public void InlineCode()
        {
            //ExStart:InlineCode
            //GistId:51b4cb9c451832f23527892e19c7bca6
            // Use a document builder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder();

            // Number of backticks is missed, one backtick will be used by default.
            Style inlineCode1BackTicks = builder.Document.Styles.Add(StyleType.Character, "InlineCode");
            builder.Font.Style = inlineCode1BackTicks;
            builder.Writeln("Text with InlineCode style with 1 backtick");

            // There will be 3 backticks.
            Style inlineCode3BackTicks = builder.Document.Styles.Add(StyleType.Character, "InlineCode.3");
            builder.Font.Style = inlineCode3BackTicks;
            builder.Writeln("Text with InlineCode style with 3 backtick");
            //ExEnd:InlineCode
        }

        [Test]
        public void Autolink()
        {
            //ExStart:Autolink
            //GistId:0697355b7f872839932388d269ed6a63
            // Use a document builder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder();

            // Insert hyperlink.
            builder.InsertHyperlink("https://www.aspose.com", "https://www.aspose.com", false);
            builder.InsertHyperlink("email@aspose.com", "mailto:email@aspose.com", false);
            //ExEnd:Autolink
        }

        [Test]
        public void Link()
        {
            //ExStart:Link
            //GistId:0697355b7f872839932388d269ed6a63
            // Use a document builder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder();

            // Insert hyperlink.
            builder.InsertHyperlink("Aspose", "https://www.aspose.com", false);
            //ExEnd:Link
        }

        [Test]
        public void Image()
        {
            //ExStart:Image
            //GistId:0697355b7f872839932388d269ed6a63
            // Use a document builder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder();

            // Insert image.
            Shape shape = builder.InsertImage(ImagesDir + "Logo.jpg");
            shape.ImageData.Title = "title";
            //ExEnd:Image
        }

        [Test]
        public void HorizontalRule()
        {
            //ExStart:HorizontalRule
            //GistId:0697355b7f872839932388d269ed6a63
            // Use a document builder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder();

            // Insert horizontal rule.
            builder.InsertHorizontalRule();
            //ExEnd:HorizontalRule
        }

        [Test]
        public void Heading()
        {
            //ExStart:Heading
            //GistId:0697355b7f872839932388d269ed6a63
            // Use a document builder to add content to the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // By default Heading styles in Word may have Bold and Italic formatting.
            //If we do not want to be emphasized, set these properties explicitly to false.
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

            // Note, emphases are also allowed inside Headings:
            builder.Font.Bold = true;
            builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
            builder.Writeln("Bold Heading1");

            doc.Save(ArtifactsDir + "WorkingWithMarkdown.Heading.md");
            //ExEnd:Heading
        }

        [Test]
        public void SetextHeading()
        {
            //ExStart:SetextHeading
            //GistId:0697355b7f872839932388d269ed6a63
            // Use a document builder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder();

            builder.ParagraphFormat.StyleName = "Heading 1";
            builder.Writeln("This is an H1 tag");

            // Reset styles from the previous paragraph to not combine styles between paragraphs.
            builder.Font.Bold = false;
            builder.Font.Italic = false;

            Style setexHeading1 = builder.Document.Styles.Add(StyleType.Paragraph, "SetextHeading1");
            builder.ParagraphFormat.Style = setexHeading1;
            builder.Document.Styles["SetextHeading1"].BaseStyleName = "Heading 1";
            builder.Writeln("Setext Heading level 1");

            builder.ParagraphFormat.Style = builder.Document.Styles["Heading 3"];
            builder.Writeln("This is an H3 tag");

            // Reset styles from the previous paragraph to not combine styles between paragraphs.
            builder.Font.Bold = false;
            builder.Font.Italic = false;

            Style setexHeading2 = builder.Document.Styles.Add(StyleType.Paragraph, "SetextHeading2");
            builder.ParagraphFormat.Style = setexHeading2;
            builder.Document.Styles["SetextHeading2"].BaseStyleName = "Heading 3";

            // Setex heading level will be reset to 2 if the base paragraph has a Heading level greater than 2.
            builder.Writeln("Setext Heading level 2");
            //ExEnd:SetextHeading

            builder.Document.Save(ArtifactsDir + "WorkingWithMarkdown.SetextHeading.md");
        }

        [Test]
        public void IndentedCode()
        {
            //ExStart:IndentedCode
            //GistId:0697355b7f872839932388d269ed6a63
            // Use a document builder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder();

            Style indentedCode = builder.Document.Styles.Add(StyleType.Paragraph, "IndentedCode");
            builder.ParagraphFormat.Style = indentedCode;
            builder.Writeln("This is an indented code");
            //ExEnd:IndentedCode
        }

        [Test]
        public void FencedCode()
        {
            //ExStart:FencedCode
            //GistId:0697355b7f872839932388d269ed6a63
            // Use a document builder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder();

            Style fencedCode = builder.Document.Styles.Add(StyleType.Paragraph, "FencedCode");
            builder.ParagraphFormat.Style = fencedCode;
            builder.Writeln("This is an fenced code");

            Style fencedCodeWithInfo = builder.Document.Styles.Add(StyleType.Paragraph, "FencedCode.C#");
            builder.ParagraphFormat.Style = fencedCodeWithInfo;
            builder.Writeln("This is a fenced code with info string");
            //ExEnd:FencedCode
        }

        [Test]
        public void Quote()
        {
            //ExStart:Quote
            //GistId:0697355b7f872839932388d269ed6a63
            // Use a document builder to add content to the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // By default a document stores blockquote style for the first level.
            builder.ParagraphFormat.StyleName = "Quote";
            builder.Writeln("Blockquote");

            // Create styles for nested levels through style inheritance.
            Style quoteLevel2 = builder.Document.Styles.Add(StyleType.Paragraph, "Quote1");
            builder.ParagraphFormat.Style = quoteLevel2;
            builder.Document.Styles["Quote1"].BaseStyleName = "Quote";
            builder.Writeln("1. Nested blockquote");

            doc.Save(ArtifactsDir + "WorkingWithMarkdown.Quote.md");
            //ExEnd:Quote
        }

        [Test]
        public void BulletedList()
        {
            //ExStart:BulletedList
            //GistId:0697355b7f872839932388d269ed6a63
            // Use a document builder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder();

            builder.ListFormat.ApplyBulletDefault();
            builder.ListFormat.List.ListLevels[0].NumberFormat = "-";

            builder.Writeln("Item 1");
            builder.Writeln("Item 2");

            builder.ListFormat.ListIndent();

            builder.Writeln("Item 2a");
            builder.Writeln("Item 2b");
            //ExEnd:BulletedList
        }

        [Test]
        public void OrderedList()
        {
            //ExStart:OrderedList
            //GistId:0697355b7f872839932388d269ed6a63
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ListFormat.ApplyNumberDefault();

            builder.Writeln("Item 1");
            builder.Writeln("Item 2");

            builder.ListFormat.ListIndent();

            builder.Writeln("Item 2a");
            builder.Writeln("Item 2b");
            //ExEnd:OrderedList
        }

        [Test]
        public void Table()
        {
            //ExStart:Table
            //GistId:0697355b7f872839932388d269ed6a63
            // Use a document builder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder();

            // Add the first row.
            builder.InsertCell();
            builder.Writeln("a");
            builder.InsertCell();
            builder.Writeln("b");

            builder.EndRow();

            // Add the second row.
            builder.InsertCell();
            builder.Writeln("c");
            builder.InsertCell();
            builder.Writeln("d");
            //ExEnd:Table
        }

        [Test]
        public void ReadMarkdownDocument()
        {
            //ExStart:ReadMarkdownDocument
            //GistId:19de942ef8827201c1dca99f76c59133
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
            //GistId:19de942ef8827201c1dca99f76c59133
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

        [Test]
        public void SupportedFeatures()
        {
            //ExStart:SupportedFeatures
            //GistId:51b4cb9c451832f23527892e19c7bca6
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Specify the "Heading 1" style for the paragraph.
            builder.InsertParagraph();
            builder.ParagraphFormat.StyleName = "Heading 1";
            builder.Write("Heading 1");

            // Specify the Italic emphasis for the paragraph.
            builder.InsertParagraph();
            // Reset styles from the previous paragraph to not combine styles between paragraphs.
            builder.ParagraphFormat.StyleName = "Normal";
            builder.Font.Italic = true;
            builder.Write("Italic Text");
            // Reset styles from the previous paragraph to not combine styles between paragraphs.
            builder.Italic = false;

            // Specify a Hyperlink for the desired text.
            builder.InsertParagraph();
            builder.InsertHyperlink("Aspose", "https://www.aspose.com", false);
            builder.Write("Aspose");

            // Save your document as a Markdown file.
            doc.Save(ArtifactsDir + "WorkingWithMarkdown.SupportedFeatures.md");
            //ExEnd:SupportedFeatures
        }
    }
}