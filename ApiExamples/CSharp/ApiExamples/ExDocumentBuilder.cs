// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections;
using System.IO;
using System.Linq;
using System.Net;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Fields;
using Aspose.Words.Tables;
using NUnit.Framework;
using Cell = Aspose.Words.Tables.Cell;
using Color = System.Drawing.Color;
using Document = Aspose.Words.Document;
using Table = Aspose.Words.Tables.Table;
using System.Drawing;
using Aspose.Words.Saving;

#if NETCOREAPP2_1 || __MOBILE__
using SkiaSharp;
#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExDocumentBuilder : ApiExampleBase
    {
        [Test]
        public void WriteAndFont()
        {
            //ExStart
            //ExFor:Font.Size
            //ExFor:Font.Bold
            //ExFor:Font.Name
            //ExFor:Font.Color
            //ExFor:Font.Underline
            //ExFor:DocumentBuilder.#ctor
            //ExSummary:Inserts formatted text using DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Specify font formatting before adding text
            Aspose.Words.Font font = builder.Font;
            font.Size = 16;
            font.Bold = true;
            font.Color = Color.Blue;
            font.Name = "Courier New";
            font.Underline = Underline.Dash;

            builder.Write("Hello world!");
            //ExEnd

            doc = DocumentHelper.SaveOpen(builder.Document);
            Run firstRun = doc.FirstSection.Body.Paragraphs[0].Runs[0];

            Assert.AreEqual("Hello world!", firstRun.GetText().Trim());
            Assert.AreEqual(16, firstRun.Font.Size);
            Assert.True(firstRun.Font.Bold);
            Assert.AreEqual("Courier New", firstRun.Font.Name);
            Assert.AreEqual(Color.Blue.ToArgb(), firstRun.Font.Color.ToArgb());
            Assert.AreEqual(Underline.Dash, firstRun.Font.Underline);
        }

        [Test]
        public void HeadersAndFooters()
        {
            //ExStart
            //ExFor:DocumentBuilder
            //ExFor:DocumentBuilder.#ctor(Document)
            //ExFor:DocumentBuilder.MoveToHeaderFooter
            //ExFor:DocumentBuilder.MoveToSection
            //ExFor:DocumentBuilder.InsertBreak
            //ExFor:DocumentBuilder.Writeln
            //ExFor:HeaderFooterType
            //ExFor:PageSetup.DifferentFirstPageHeaderFooter
            //ExFor:PageSetup.OddAndEvenPagesHeaderFooter
            //ExFor:BreakType
            //ExSummary:Shows how to create headers and footers in a document using DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Specify that we want headers and footers different for first, even and odd pages
            builder.PageSetup.DifferentFirstPageHeaderFooter = true;
            builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

            // Create the headers
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.Write("Header for the first page");
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
            builder.Write("Header for even pages");
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Header for all other pages");

            // Create three pages in the document
            builder.MoveToSection(0);
            builder.Writeln("Page1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page3");

            doc.Save(ArtifactsDir + "DocumentBuilder.HeadersAndFooters.docx");
            //ExEnd

            HeaderFooterCollection headersFooters = 
                new Document(ArtifactsDir + "DocumentBuilder.HeadersAndFooters.docx").FirstSection.HeadersFooters;

            Assert.AreEqual(3, headersFooters.Count);
            Assert.AreEqual("Header for the first page", headersFooters[HeaderFooterType.HeaderFirst].GetText().Trim());
            Assert.AreEqual("Header for even pages", headersFooters[HeaderFooterType.HeaderEven].GetText().Trim());
            Assert.AreEqual("Header for all other pages", headersFooters[HeaderFooterType.HeaderPrimary].GetText().Trim());

        }

        [Test]
        public void MergeFields()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertField(String)
            //ExFor:DocumentBuilder.MoveToMergeField(String, Boolean, Boolean)
            //ExSummary:Shows how to insert merge fields and move between them.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField(@"MERGEFIELD MyMergeField1 \* MERGEFORMAT");
            builder.InsertField(@"MERGEFIELD MyMergeField2 \* MERGEFORMAT");

            // The second merge field starts immediately after the end of the first
            // We'll move the builder's cursor to the end of the first so we can split them by text
            builder.MoveToMergeField("MyMergeField1", true, false);
            Assert.AreEqual(doc.Range.Fields[1].Start, builder.CurrentNode);

            builder.Write(" Text between our two merge fields. ");

            doc.Save(ArtifactsDir + "DocumentBuilder.MergeFields.docx");
            //ExEnd		

            doc = new Document(ArtifactsDir + "DocumentBuilder.MergeFields.docx");

            Assert.AreEqual(2, doc.Range.Fields.Count);

            TestUtil.VerifyField(FieldType.FieldMergeField, @"MERGEFIELD MyMergeField1 \* MERGEFORMAT", "«MyMergeField1»", doc.Range.Fields[0]);
            TestUtil.VerifyField(FieldType.FieldMergeField, @"MERGEFIELD MyMergeField2 \* MERGEFORMAT", "«MyMergeField2»", doc.Range.Fields[1]);
        }

        [Test]
        public void InsertHorizontalRule()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertHorizontalRule
            //ExFor:ShapeBase.IsHorizontalRule
            //ExFor:Shape.HorizontalRuleFormat
            //ExFor:HorizontalRuleFormat
            //ExFor:HorizontalRuleFormat.Alignment
            //ExFor:HorizontalRuleFormat.WidthPercent
            //ExFor:HorizontalRuleFormat.Height
            //ExFor:HorizontalRuleFormat.Color
            //ExFor:HorizontalRuleFormat.NoShade
            //ExSummary:Shows how to insert horizontal rule shape in a document and customize the formatting.
            // Use a document builder to insert a horizontal rule
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            Shape shape = builder.InsertHorizontalRule();

            HorizontalRuleFormat horizontalRuleFormat = shape.HorizontalRuleFormat;
            horizontalRuleFormat.Alignment = HorizontalRuleAlignment.Center;
            horizontalRuleFormat.WidthPercent = 70;
            horizontalRuleFormat.Height = 3;
            horizontalRuleFormat.Color = Color.Blue;
            horizontalRuleFormat.NoShade = true;

            Assert.True(shape.IsHorizontalRule);
            Assert.True(shape.HorizontalRuleFormat.NoShade);
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual(HorizontalRuleAlignment.Center, shape.HorizontalRuleFormat.Alignment);
            Assert.AreEqual(70, shape.HorizontalRuleFormat.WidthPercent);
            Assert.AreEqual(3, shape.HorizontalRuleFormat.Height);
            Assert.AreEqual(Color.Blue.ToArgb(), shape.HorizontalRuleFormat.Color.ToArgb());
        }

        [Test(Description = "Checking the boundary conditions of WidthPercent and Height properties")]
        public void HorizontalRuleFormatExceptions()
        {
            DocumentBuilder builder = new DocumentBuilder();
            Shape shape = builder.InsertHorizontalRule();

            HorizontalRuleFormat horizontalRuleFormat = shape.HorizontalRuleFormat;
            horizontalRuleFormat.WidthPercent = 1;
            horizontalRuleFormat.WidthPercent = 100;
            Assert.That(() => horizontalRuleFormat.WidthPercent = 0, Throws.TypeOf<ArgumentOutOfRangeException>());
            Assert.That(() => horizontalRuleFormat.WidthPercent = 101, Throws.TypeOf<ArgumentOutOfRangeException>());
            
            horizontalRuleFormat.Height = 0;
            horizontalRuleFormat.Height = 1584;
            Assert.That(() => horizontalRuleFormat.Height = -1, Throws.TypeOf<ArgumentOutOfRangeException>());
            Assert.That(() => horizontalRuleFormat.Height = 1585, Throws.TypeOf<ArgumentOutOfRangeException>());
        }

        [Test]
        public void InsertHyperlink()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertHyperlink
            //ExFor:Font.ClearFormatting
            //ExFor:Font.Color
            //ExFor:Font.Underline
            //ExFor:Underline
            //ExSummary:Shows how to insert a hyperlink into a document using DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Please make sure to visit ");

            // Specify font formatting for the hyperlink
            builder.Font.Color = Color.Blue;
            builder.Font.Underline = Underline.Single;

            // Insert the link
            builder.InsertHyperlink("Aspose Website", "http://www.aspose.com", false);

            // Revert to default formatting
            builder.Font.ClearFormatting();
            builder.Write(" for more information.");

            // Holding Ctrl and left clicking on the field in Microsoft Word will take you to the link's address in a web browser
            doc.Save(ArtifactsDir + "DocumentBuilder.InsertHyperlink.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertHyperlink.docx");

            FieldHyperlink hyperlink = (FieldHyperlink)doc.Range.Fields[0];
            TestUtil.VerifyWebResponseStatusCode(HttpStatusCode.OK, hyperlink.Address);

            Run fieldContents = (Run)hyperlink.Start.NextSibling;

            Assert.AreEqual(Color.Blue.ToArgb(), fieldContents.Font.Color.ToArgb());
            Assert.AreEqual(Underline.Single, fieldContents.Font.Underline);
            Assert.AreEqual("HYPERLINK \"http://www.aspose.com\"", fieldContents.GetText().Trim());
        }

        [Test]
        public void PushPopFont()
        {
            //ExStart
            //ExFor:DocumentBuilder.PushFont
            //ExFor:DocumentBuilder.PopFont
            //ExFor:DocumentBuilder.InsertHyperlink
            //ExSummary:Shows how to use temporarily save and restore character formatting when building a document with DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set up font formatting and write text that goes before the hyperlink
            builder.Font.Name = "Arial";
            builder.Font.Size = 24;
            builder.Font.Bold = true;
            builder.Write("To visit Google, hold Ctrl and click ");

            // Save the font formatting so we use different formatting for hyperlink and restore old formatting later
            builder.PushFont();

            // Set new font formatting for the hyperlink and insert the hyperlink
            // The "Hyperlink" style is a Microsoft Word built-in style so we don't have to worry to 
            // create it, it will be created automatically if it does not yet exist in the document
            builder.Font.StyleIdentifier = StyleIdentifier.Hyperlink;
            builder.InsertHyperlink("here", "http://www.google.com", false);

            // Restore the formatting that was before the hyperlink
            builder.PopFont();

            builder.Write(". We hope you enjoyed the example.");

            doc.Save(ArtifactsDir + "DocumentBuilder.PushPopFont.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.PushPopFont.docx");
            RunCollection runs = doc.FirstSection.Body.FirstParagraph.Runs;

            Assert.AreEqual(4, runs.Count);

            Assert.AreEqual("To visit Google, hold Ctrl and click", runs[0].GetText().Trim());
            Assert.AreEqual(". We hope you enjoyed the example.", runs[3].GetText().Trim());
            Assert.AreEqual(runs[0].Font.Color, runs[3].Font.Color);
            Assert.AreEqual(runs[0].Font.Underline, runs[3].Font.Underline);

            Assert.AreEqual("here", runs[2].GetText().Trim());
            Assert.AreEqual(Color.Blue.ToArgb(), runs[2].Font.Color.ToArgb());
            Assert.AreEqual(Underline.Single, runs[2].Font.Underline);
            Assert.AreNotEqual(runs[0].Font.Color, runs[2].Font.Color);
            Assert.AreNotEqual(runs[0].Font.Underline, runs[2].Font.Underline);
        }

#if NET462 || JAVA
        [Test]
        public void InsertWatermark()
        {
            //ExStart
            //ExFor:DocumentBuilder.MoveToHeaderFooter
            //ExFor:PageSetup.PageWidth
            //ExFor:PageSetup.PageHeight
            //ExFor:DocumentBuilder.InsertImage(Image)
            //ExFor:WrapType
            //ExFor:RelativeHorizontalPosition
            //ExFor:RelativeVerticalPosition
            //ExSummary:Shows how to a watermark image into a document using DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The best place for the watermark image is in the header or footer so it is shown on every page
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            Image image = Image.FromFile(ImageDir + "Transparent background logo.png");

            // Insert a floating picture
            Shape shape = builder.InsertImage(image);
            shape.WrapType = WrapType.None;
            shape.BehindText = true;

            shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;

            // Calculate image left and top position so it appears in the center of the page
            shape.Left = (builder.PageSetup.PageWidth - shape.Width) / 2;
            shape.Top = (builder.PageSetup.PageHeight - shape.Height) / 2;

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertWatermark.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertWatermark.docx");
            shape = (Shape)doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Png, shape);
            Assert.AreEqual(WrapType.None, shape.WrapType);
            Assert.True(shape.BehindText);
            Assert.AreEqual(RelativeHorizontalPosition.Page, shape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Page, shape.RelativeVerticalPosition);
            Assert.AreEqual((doc.FirstSection.PageSetup.PageWidth - shape.Width) / 2, shape.Left);
            Assert.AreEqual((doc.FirstSection.PageSetup.PageHeight - shape.Height) / 2, shape.Top);
        }

        [Test]
        public void InsertOleObject()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertOleObject(String, Boolean, Boolean, Image)
            //ExFor:DocumentBuilder.InsertOleObject(String, String, Boolean, Boolean, Image)
            //ExFor:DocumentBuilder.InsertOleObjectAsIcon(String, Boolean, String, String)
            //ExSummary:Shows how to insert an OLE object into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            // Insert ole object
            Image representingImage = Image.FromFile(ImageDir + "Logo.jpg");
            builder.InsertOleObject(MyDir + "Spreadsheet.xlsx", false, false, representingImage);

            // Insert ole object with ProgId
            builder.InsertOleObject(MyDir + "Spreadsheet.xlsx", "Excel.Sheet", false, true, null);

            // Insert ole object as Icon
            // There is one limitation for now: the maximum size of the icon must be 32x32 for the correct display
            builder.InsertOleObjectAsIcon(MyDir + "Presentation.pptx", false, ImageDir + "Logo icon.ico",
                "Caption (can not be null)");

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertOleObject.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertOleObject.docx");
            Shape shape = (Shape)doc.GetChild(NodeType.Shape,0, true);
            
            Assert.AreEqual(ShapeType.OleObject, shape.ShapeType);
            Assert.AreEqual("Excel.Sheet.12", shape.OleFormat.ProgId);
            Assert.AreEqual(".xlsx", shape.OleFormat.SuggestedExtension);

            shape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            Assert.AreEqual(ShapeType.OleObject, shape.ShapeType);
            Assert.AreEqual("Package", shape.OleFormat.ProgId);
            Assert.AreEqual(".xlsx", shape.OleFormat.SuggestedExtension);

            shape = (Shape)doc.GetChild(NodeType.Shape, 2, true);

            Assert.AreEqual(ShapeType.OleObject, shape.ShapeType);
            Assert.AreEqual("PowerPoint.Show.12", shape.OleFormat.ProgId);
            Assert.AreEqual(".pptx", shape.OleFormat.SuggestedExtension);
        }
#elif NETCOREAPP2_1 || __MOBILE__
        [Test]
        public void InsertWatermarkNetStandard2()
        {
            //ExStart
            //ExFor:DocumentBuilder.MoveToHeaderFooter
            //ExFor:PageSetup.PageWidth
            //ExFor:PageSetup.PageHeight
            //ExFor:DocumentBuilder.InsertImage(Image)
            //ExFor:WrapType
            //ExFor:RelativeHorizontalPosition
            //ExFor:RelativeVerticalPosition
            //ExSummary:Shows how to insert a watermark image into a document using DocumentBuilder (.NetStandard 2.0).
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The best place for the watermark image is in the header or footer so it is shown on every page
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            Shape shape;
            using (SKBitmap image = SKBitmap.Decode(ImageDir + "Transparent background logo.png"))
            {
                // Insert a floating picture
                shape = builder.InsertImage(image);
                shape.WrapType = WrapType.None;
                shape.BehindText = true;

                shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
                shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;

                // Calculate image left and top position so it appears in the center of the page
                shape.Left = (builder.PageSetup.PageWidth - shape.Width) / 2;
                shape.Top = (builder.PageSetup.PageHeight - shape.Height) / 2;
            }

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertWatermarkNetStandard2.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertWatermarkNetStandard2.docx");
            shape = (Shape)doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Png, shape);
            Assert.AreEqual(WrapType.None, shape.WrapType);
            Assert.True(shape.BehindText);
            Assert.AreEqual(RelativeHorizontalPosition.Page, shape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Page, shape.RelativeVerticalPosition);
            Assert.AreEqual((doc.FirstSection.PageSetup.PageWidth - shape.Width) / 2, shape.Left);
            Assert.AreEqual((doc.FirstSection.PageSetup.PageHeight - shape.Height) / 2, shape.Top);
        }

        [Test]
        public void InsertOleObjectNetStandard2()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertOleObject(String, Boolean, Boolean, Image)
            //ExFor:DocumentBuilder.InsertOleObject(String, String, Boolean, Boolean, Image)
            //ExSummary:Shows how to insert an OLE object into a document (.NetStandard 2.0).
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            using (SKBitmap representingImage = SKBitmap.Decode(ImageDir + "Logo.jpg"))
            {
                // OleObject
                builder.InsertOleObject(MyDir + "Spreadsheet.xlsx", false, false, representingImage);
                // OleObject with ProgId
                builder.InsertOleObject(MyDir + "Spreadsheet.xlsx", "Excel.Sheet", false, false,
                    representingImage);
            }

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertOleObjectNetStandard2.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertOleObjectNetStandard2.docx");
            Shape shape = (Shape)doc.GetChild(NodeType.Shape,0, true);
            
            Assert.AreEqual(ShapeType.OleObject, shape.ShapeType);
            Assert.AreEqual("Excel.Sheet.12", shape.OleFormat.ProgId);
            Assert.AreEqual(".xlsx", shape.OleFormat.SuggestedExtension);

            shape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            Assert.AreEqual(ShapeType.OleObject, shape.ShapeType);
            Assert.AreEqual("Package", shape.OleFormat.ProgId);
            Assert.AreEqual(".xlsx", shape.OleFormat.SuggestedExtension);
        }
#endif

        [Test]
        public void InsertHtml()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertHtml(String)
            //ExSummary:Shows how to insert Html content into a document using a builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            const string html = "<P align='right'>Paragraph right</P>" + "<b>Implicit paragraph left</b>" +
                                "<div align='center'>Div center</div>" + "<h1 align='left'>Heading 1 left.</h1>";

            builder.InsertHtml(html);

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertHtml.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertHtml.docx");
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            Assert.AreEqual("Paragraph right", paragraphs[0].GetText().Trim());
            Assert.AreEqual(ParagraphAlignment.Right, paragraphs[0].ParagraphFormat.Alignment);

            Assert.AreEqual("Implicit paragraph left", paragraphs[1].GetText().Trim());
            Assert.AreEqual(ParagraphAlignment.Left, paragraphs[1].ParagraphFormat.Alignment);
            Assert.True(paragraphs[1].Runs[0].Font.Bold);

            Assert.AreEqual("Div center", paragraphs[2].GetText().Trim());
            Assert.AreEqual(ParagraphAlignment.Center, paragraphs[2].ParagraphFormat.Alignment);

            Assert.AreEqual("Heading 1 left.", paragraphs[3].GetText().Trim());
            Assert.AreEqual("Heading 1", paragraphs[3].ParagraphFormat.Style.Name);
        }

        [Test]
        public void InsertHtmlWithFormatting()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertHtml(String, Boolean)
            //ExSummary:Shows how to insert Html content into a document using a builder while applying the builder's formatting. 
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set the builder's text alignment
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Distributed;

            // If we insert text while setting useBuilderFormatting to true, any formatting applied to the builder will be applied to inserted .html content
            // However, if the html text has formatting coded into it, that formatting takes precedence over the builder's formatting
            // In this case, elements with "align" attributes do not get affected by the ParagraphAlignment we specified above
            builder.InsertHtml(
                "<P align='right'>Paragraph right</P>" + "<b>Implicit paragraph left</b>" +
                "<div align='center'>Div center</div>" + "<h1 align='left'>Heading 1 left.</h1>", true);

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertHtmlWithFormatting.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertHtmlWithFormatting.docx");
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            Assert.AreEqual("Paragraph right", paragraphs[0].GetText().Trim());
            Assert.AreEqual(ParagraphAlignment.Right, paragraphs[0].ParagraphFormat.Alignment);

            Assert.AreEqual("Implicit paragraph left", paragraphs[1].GetText().Trim());
            Assert.AreEqual(ParagraphAlignment.Distributed, paragraphs[1].ParagraphFormat.Alignment);
            Assert.True(paragraphs[1].Runs[0].Font.Bold);

            Assert.AreEqual("Div center", paragraphs[2].GetText().Trim());
            Assert.AreEqual(ParagraphAlignment.Center, paragraphs[2].ParagraphFormat.Alignment);

            Assert.AreEqual("Heading 1 left.", paragraphs[3].GetText().Trim());
            Assert.AreEqual("Heading 1", paragraphs[3].ParagraphFormat.Style.Name);
        }

        [Test]
        public void MathML()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            const string mathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><msub><mi>a</mi><mrow><mn>1</mn></mrow></msub><mo>+</mo><msub><mi>b</mi><mrow><mn>1</mn></mrow></msub></mrow></math>";

            builder.InsertHtml(mathMl);

            doc.Save(ArtifactsDir + "DocumentBuilder.MathML.docx");
            doc.Save(ArtifactsDir + "DocumentBuilder.MathML.pdf");

            Assert.IsTrue(DocumentHelper.CompareDocs(GoldsDir + "DocumentBuilder.MathML Gold.docx", ArtifactsDir + "DocumentBuilder.MathML.docx"));
        }

        [Test]
        public void InsertTextAndBookmark()
        {
            //ExStart
            //ExFor:DocumentBuilder.StartBookmark
            //ExFor:DocumentBuilder.EndBookmark
            //ExSummary:Shows how to add some text into the document and encloses the text in a bookmark using DocumentBuilder.
            DocumentBuilder builder = new DocumentBuilder();

            builder.StartBookmark("MyBookmark");
            builder.Writeln("Text inside a bookmark.");
            builder.EndBookmark("MyBookmark");
            //ExEnd

            Document doc = DocumentHelper.SaveOpen(builder.Document);

            Assert.AreEqual(1, doc.Range.Bookmarks.Count);
            Assert.AreEqual("MyBookmark", doc.Range.Bookmarks[0].Name);
            Assert.AreEqual("Text inside a bookmark.", doc.Range.Bookmarks[0].Text.Trim());
        }

        [Test]
        public void CreateForm()
        {
            //ExStart
            //ExFor:TextFormFieldType
            //ExFor:DocumentBuilder.InsertTextInput
            //ExFor:DocumentBuilder.InsertComboBox
            //ExSummary:Shows how to build a form field.
            DocumentBuilder builder = new DocumentBuilder();

            // Insert a text form field for input a name
            builder.InsertTextInput("", TextFormFieldType.Regular, "", "Enter your name here", 30);

            // Insert two blank lines
            builder.Writeln("");
            builder.Writeln("");

            string[] items =
            {
                "-- Select your favorite footwear --", "Sneakers", "Oxfords", "Flip-flops", "Other",
                "I prefer to be barefoot"
            };

            // Insert a combo box to select a footwear type
            builder.InsertComboBox("", items, 0);

            // Insert 2 blank lines
            builder.Writeln("");
            builder.Writeln("");

            builder.Document.Save(ArtifactsDir + "DocumentBuilder.CreateForm.docx");
            //ExEnd

            Document doc = new Document(ArtifactsDir + "DocumentBuilder.CreateForm.docx");
            FormField formField = doc.Range.FormFields[0];

            Assert.AreEqual(TextFormFieldType.Regular, formField.TextInputType);
            Assert.AreEqual("Enter your name here", formField.Result);

            formField = doc.Range.FormFields[1];

            Assert.AreEqual(TextFormFieldType.Regular, formField.TextInputType);
            Assert.AreEqual("-- Select your favorite footwear --", formField.Result);
            Assert.AreEqual(0, formField.DropDownSelectedIndex);
            Assert.AreEqual(new[] { "-- Select your favorite footwear --", "Sneakers", "Oxfords", "Flip-flops", "Other",
                "I prefer to be barefoot" }, formField.DropDownItems.ToArray());
        }

        [Test]
        public void InsertCheckBox()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertCheckBox(string, bool, bool, int)
            //ExFor:DocumentBuilder.InsertCheckBox(String, bool, int)
            //ExSummary:Shows how to insert checkboxes to the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertCheckBox(string.Empty, false, false, 0);
            builder.InsertCheckBox("CheckBox_Default", true, true, 50);
            builder.InsertCheckBox("CheckBox_OnlyCheckedValue", true, 100);
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);

            // Get checkboxes from the document
            FormFieldCollection formFields = doc.Range.FormFields;

            // Check that is the right checkbox
            Assert.AreEqual(string.Empty, formFields[0].Name);

            //Assert that parameters sets correctly
            Assert.AreEqual(false, formFields[0].Checked);
            Assert.AreEqual(false, formFields[0].Default);
            Assert.AreEqual(10, formFields[0].CheckBoxSize);

            // Check that is the right checkbox
            // Please pay attention that MS Word allows strings with at most 20 characters
            Assert.AreEqual("CheckBox_Default", formFields[1].Name);

            //Assert that parameters sets correctly
            Assert.AreEqual(true, formFields[1].Checked);
            Assert.AreEqual(true, formFields[1].Default);
            Assert.AreEqual(50, formFields[1].CheckBoxSize);

            // Check that is the right checkbox
            // Please pay attention that MS Word allows strings with at most 20 characters
            Assert.AreEqual("CheckBox_OnlyChecked", formFields[2].Name);

            // Assert that parameters sets correctly
            Assert.AreEqual(true, formFields[2].Checked);
            Assert.AreEqual(true, formFields[2].Default);
            Assert.AreEqual(100, formFields[2].CheckBoxSize);
        }

        [Test]
        public void InsertCheckBoxEmptyName()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Checking that the checkbox insertion with an empty name working correctly
            builder.InsertCheckBox("", true, false, 1);
            builder.InsertCheckBox(string.Empty, false, 1);
        }

        [Test]
        public void WorkingWithNodes()
        {
            //ExStart
            //ExFor:DocumentBuilder.MoveTo(Node)
            //ExFor:DocumentBuilder.MoveToBookmark(String)
            //ExFor:DocumentBuilder.CurrentParagraph
            //ExFor:DocumentBuilder.CurrentNode
            //ExFor:DocumentBuilder.MoveToDocumentStart
            //ExFor:DocumentBuilder.MoveToDocumentEnd
            //ExFor:DocumentBuilder.IsAtEndOfParagraph
            //ExFor:DocumentBuilder.IsAtStartOfParagraph
            //ExSummary:Shows how to move a DocumentBuilder to different nodes in a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a bookmark and add content to it using a DocumentBuilder
            builder.StartBookmark("MyBookmark");
            builder.Writeln("Bookmark contents.");
            builder.EndBookmark("MyBookmark");

            // The node that the DocumentBuilder is currently at is past the boundaries of the bookmark  
            Assert.AreEqual(doc.Range.Bookmarks[0].BookmarkEnd, builder.CurrentParagraph.FirstChild);

            // If we wish to revise the content of our bookmark with the DocumentBuilder, we can move back to it like this
            builder.MoveToBookmark("MyBookmark");

            // Now we're located between the bookmark's BookmarkStart and BookmarkEnd nodes, so any text the builder adds will be within it
            Assert.AreEqual(doc.Range.Bookmarks[0].BookmarkStart, builder.CurrentParagraph.FirstChild);

            // We can move the builder to an individual node,
            // which in this case will be the first node of the first paragraph, like this
            builder.MoveTo(doc.FirstSection.Body.FirstParagraph.GetChildNodes(NodeType.Any, false)[0]);

            Assert.AreEqual(NodeType.BookmarkStart, builder.CurrentNode.NodeType);
            Assert.IsTrue(builder.IsAtStartOfParagraph);

            // A shorter way of moving the very start/end of a document is with these methods
            builder.MoveToDocumentEnd();

            Assert.IsTrue(builder.IsAtEndOfParagraph);

            builder.MoveToDocumentStart();

            Assert.IsTrue(builder.IsAtStartOfParagraph);
            //ExEnd
        }

        [Test]
        public void FillMergeFields()
        {
            //ExStart
            //ExFor:DocumentBuilder.MoveToMergeField(String)
            //ExFor:DocumentBuilder.Bold
            //ExFor:DocumentBuilder.Italic
            //ExSummary:Shows how to fill MERGEFIELDs with data with a DocumentBuilder and without a mail merge.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert some MERGEFIELDS, which accept data from columns of the same name in a data source during a mail merge
            builder.InsertField(" MERGEFIELD Chairman ");
            builder.InsertField(" MERGEFIELD ChiefFinancialOfficer ");
            builder.InsertField(" MERGEFIELD ChiefTechnologyOfficer ");

            // They can also be filled in manually like this
            builder.MoveToMergeField("Chairman");
            builder.Bold = true;
            builder.Writeln("John Doe");

            builder.MoveToMergeField("ChiefFinancialOfficer");
            builder.Italic = true;
            builder.Writeln("Jane Doe");

            builder.MoveToMergeField("ChiefTechnologyOfficer");
            builder.Italic = true;
            builder.Writeln("John Bloggs");

            doc.Save(ArtifactsDir + "DocumentBuilder.FillMergeFields.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.FillMergeFields.docx");
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            Assert.True(paragraphs[0].Runs[0].Font.Bold);
            Assert.AreEqual("John Doe", paragraphs[0].Runs[0].GetText().Trim());

            Assert.True(paragraphs[1].Runs[0].Font.Italic);
            Assert.AreEqual("Jane Doe", paragraphs[1].Runs[0].GetText().Trim());

            Assert.True(paragraphs[2].Runs[0].Font.Italic);
            Assert.AreEqual("John Bloggs", paragraphs[2].Runs[0].GetText().Trim());

        }

        [Test]
        public void InsertToc()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertTableOfContents
            //ExFor:Document.UpdateFields
            //ExFor:DocumentBuilder.#ctor(Document)
            //ExFor:ParagraphFormat.StyleIdentifier
            //ExFor:DocumentBuilder.InsertBreak
            //ExFor:BreakType
            //ExSummary:Shows how to insert a Table of contents (TOC) into a document using heading styles as entries.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table of contents at the beginning of the document,
            // and set it to pick up paragraphs with headings of levels 1 to 3 and entries to act like hyperlinks
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

            // Start the actual document content on the second page
            builder.InsertBreak(BreakType.PageBreak);

            // Build a document with complex structure by applying different heading styles thus creating TOC entries
            // The heading levels we use below will affect the list levels in which these items will appear in the TOC,
            // and only levels 1-3 will be picked up by our TOC due to its settings
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

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
            builder.Writeln("Heading 3.1.1");
            builder.Writeln("Heading 3.1.2");
            builder.Writeln("Heading 3.1.3");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Heading 3.2");
            builder.Writeln("Heading 3.3");

            // Call the method below to update the TOC and save
            doc.UpdateFields();
            doc.Save(ArtifactsDir + "DocumentBuilder.InsertToc.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertToc.docx");
            FieldToc tableOfContents = (FieldToc)doc.Range.Fields[0];

            Assert.AreEqual("1-3", tableOfContents.HeadingLevelRange);
            Assert.IsTrue(tableOfContents.InsertHyperlinks);
            Assert.IsTrue(tableOfContents.HideInWebLayout);
            Assert.IsTrue(tableOfContents.UseParagraphOutlineLevel);
        }

        [Test]
        public void InsertTable()
        {
            //ExStart
            //ExFor:DocumentBuilder
            //ExFor:DocumentBuilder.Write
            //ExFor:DocumentBuilder.StartTable
            //ExFor:DocumentBuilder.InsertCell
            //ExFor:DocumentBuilder.EndRow
            //ExFor:DocumentBuilder.EndTable
            //ExFor:DocumentBuilder.CellFormat
            //ExFor:DocumentBuilder.RowFormat
            //ExFor:CellFormat
            //ExFor:CellFormat.FitText
            //ExFor:CellFormat.Width
            //ExFor:CellFormat.VerticalAlignment
            //ExFor:CellFormat.Shading
            //ExFor:CellFormat.Orientation
            //ExFor:CellFormat.WrapText
            //ExFor:RowFormat
            //ExFor:RowFormat.Borders
            //ExFor:RowFormat.ClearFormatting
            //ExFor:Shading.ClearFormatting
            //ExSummary:Shows how to build a nice bordered table.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building a table
            builder.StartTable();
            
            // Set the appropriate paragraph, cell, and row formatting. The formatting properties are preserved
            // until they are explicitly modified so there's no need to set them for each row or cell
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            builder.CellFormat.ClearFormatting();
            builder.CellFormat.Width = 150;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.GreenYellow;
            builder.CellFormat.WrapText = false;
            builder.CellFormat.FitText = true;

            builder.RowFormat.ClearFormatting();
            builder.RowFormat.HeightRule = HeightRule.Exactly;
            builder.RowFormat.Height = 50;
            builder.RowFormat.Borders.LineStyle = LineStyle.Engrave3D;
            builder.RowFormat.Borders.Color = Color.Orange;

            builder.InsertCell();
            builder.Write("Row 1, Col 1");

            builder.InsertCell();
            builder.Write("Row 1, Col 2");

            builder.EndRow();

            // Remove the shading (clear background)
            builder.CellFormat.Shading.ClearFormatting();

            builder.InsertCell();
            builder.Write("Row 2, Col 1");

            builder.InsertCell();
            builder.Write("Row 2, Col 2");

            builder.EndRow();

            builder.InsertCell();

            // Make the row height bigger so that a vertically oriented text could fit into cells
            builder.RowFormat.Height = 150;
            builder.CellFormat.Orientation = TextOrientation.Upward;
            builder.Write("Row 3, Col 1");

            builder.InsertCell();
            builder.CellFormat.Orientation = TextOrientation.Downward;
            builder.Write("Row 3, Col 2");

            builder.EndRow();

            builder.EndTable();

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertTable.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertTable.docx");
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.AreEqual("Row 1, Col 1\a", table.Rows[0].Cells[0].GetText().Trim());
            Assert.AreEqual("Row 1, Col 2\a", table.Rows[0].Cells[1].GetText().Trim());
            Assert.AreEqual(HeightRule.Exactly, table.Rows[0].RowFormat.HeightRule);
            Assert.AreEqual(50.0d, table.Rows[0].RowFormat.Height);
            Assert.AreEqual(LineStyle.Engrave3D, table.Rows[0].RowFormat.Borders.LineStyle);
            Assert.AreEqual(Color.Orange.ToArgb(), table.Rows[0].RowFormat.Borders.Color.ToArgb());

            foreach (Cell c in table.Rows[0].Cells)
            {
                Assert.AreEqual(150, c.CellFormat.Width);
                Assert.AreEqual(CellVerticalAlignment.Center, c.CellFormat.VerticalAlignment);
                Assert.AreEqual(Color.GreenYellow.ToArgb(), c.CellFormat.Shading.BackgroundPatternColor.ToArgb());
                Assert.IsFalse(c.CellFormat.WrapText);
                Assert.IsTrue(c.CellFormat.FitText);

                Assert.AreEqual(ParagraphAlignment.Center, c.FirstParagraph.ParagraphFormat.Alignment);
            }

            Assert.AreEqual("Row 2, Col 1\a", table.Rows[1].Cells[0].GetText().Trim());
            Assert.AreEqual("Row 2, Col 2\a", table.Rows[1].Cells[1].GetText().Trim());


            foreach (Cell c in table.Rows[1].Cells)
            {
                Assert.AreEqual(150, c.CellFormat.Width);
                Assert.AreEqual(CellVerticalAlignment.Center, c.CellFormat.VerticalAlignment);
                Assert.AreEqual(Color.Empty.ToArgb(), c.CellFormat.Shading.BackgroundPatternColor.ToArgb());
                Assert.IsFalse(c.CellFormat.WrapText);
                Assert.IsTrue(c.CellFormat.FitText);

                Assert.AreEqual(ParagraphAlignment.Center, c.FirstParagraph.ParagraphFormat.Alignment);
            }

            Assert.AreEqual(150, table.Rows[2].RowFormat.Height);

            Assert.AreEqual("Row 3, Col 1\a", table.Rows[2].Cells[0].GetText().Trim());
            Assert.AreEqual(TextOrientation.Upward, table.Rows[2].Cells[0].CellFormat.Orientation);
            Assert.AreEqual(ParagraphAlignment.Center, table.Rows[2].Cells[0].FirstParagraph.ParagraphFormat.Alignment);

            Assert.AreEqual("Row 3, Col 2\a", table.Rows[2].Cells[1].GetText().Trim());
            Assert.AreEqual(TextOrientation.Downward, table.Rows[2].Cells[1].CellFormat.Orientation);
            Assert.AreEqual(ParagraphAlignment.Center, table.Rows[2].Cells[1].FirstParagraph.ParagraphFormat.Alignment);
        }

        [Test]
        public void InsertTableWithStyle()
        {
            //ExStart
            //ExFor:Table.StyleIdentifier
            //ExFor:Table.StyleOptions
            //ExFor:TableStyleOptions
            //ExFor:Table.AutoFit
            //ExFor:AutoFitBehavior
            //ExSummary:Shows how to build a new table with a table style applied.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            Table table = builder.StartTable();

            // We must insert at least one row first before setting any table formatting
            builder.InsertCell();

            // Set the table style used based of the unique style identifier
            // Note that not all table styles are available when saving as .doc format
            table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

            // Apply which features should be formatted by the style
            table.StyleOptions =
                TableStyleOptions.FirstColumn | TableStyleOptions.RowBands | TableStyleOptions.FirstRow;
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Continue with building the table as normal
            builder.Writeln("Item");
            builder.CellFormat.RightPadding = 40;
            builder.InsertCell();
            builder.Writeln("Quantity (kg)");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln("Apples");
            builder.InsertCell();
            builder.Writeln("20");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln("Bananas");
            builder.InsertCell();
            builder.Writeln("40");
            builder.EndRow();

            builder.InsertCell();
            builder.Writeln("Carrots");
            builder.InsertCell();
            builder.Writeln("50");
            builder.EndRow();

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertTableWithStyle.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertTableWithStyle.docx");

            // Verify that the style was set by expanding to direct formatting
            doc.ExpandTableStylesToDirectFormatting();

            Assert.AreEqual("Medium Shading 1 Accent 1", table.Style.Name);
            Assert.AreEqual(TableStyleOptions.FirstColumn | TableStyleOptions.RowBands | TableStyleOptions.FirstRow,
                table.StyleOptions);
            Assert.AreEqual(189, table.FirstRow.FirstCell.CellFormat.Shading.BackgroundPatternColor.B);
            Assert.AreEqual(Color.White.ToArgb(), table.FirstRow.FirstCell.FirstParagraph.Runs[0].Font.Color.ToArgb());
            Assert.AreNotEqual(Color.LightBlue.ToArgb(),
                table.LastRow.FirstCell.CellFormat.Shading.BackgroundPatternColor.B);
            Assert.AreEqual(Color.Empty.ToArgb(), table.LastRow.FirstCell.FirstParagraph.Runs[0].Font.Color.ToArgb());
        }

        [Test]
        public void InsertTableSetHeadingRow()
        {
            //ExStart
            //ExFor:RowFormat.HeadingFormat
            //ExSummary:Shows how to build a table which include heading rows that repeat on subsequent pages. 
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartTable();
            builder.RowFormat.HeadingFormat = true;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.Width = 100;
            builder.InsertCell();
            builder.Writeln("Heading row 1");
            builder.EndRow();
            builder.InsertCell();
            builder.Writeln("Heading row 2");
            builder.EndRow();

            builder.CellFormat.Width = 50;
            builder.ParagraphFormat.ClearFormatting();

            // Insert some content so the table is long enough to continue onto the next page
            for (int i = 0; i < 50; i++)
            {
                builder.InsertCell();
                builder.RowFormat.HeadingFormat = false;
                builder.Write("Column 1 Text");
                builder.InsertCell();
                builder.Write("Column 2 Text");
                builder.EndRow();
            }

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertTableSetHeadingRow.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertTableSetHeadingRow.docx");
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.True(table.FirstRow.RowFormat.HeadingFormat);
            Assert.True(table.Rows[1].RowFormat.HeadingFormat);
            Assert.False(table.Rows[2].RowFormat.HeadingFormat);
        }

        [Test]
        public void InsertTableWithPreferredWidth()
        {
            //ExStart
            //ExFor:Table.PreferredWidth
            //ExFor:PreferredWidth.FromPercent
            //ExFor:PreferredWidth
            //ExSummary:Shows how to set a table to auto fit to 50% of the page width.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table with a width that takes up half the page width
            Table table = builder.StartTable();

            // Insert a few cells
            builder.InsertCell();
            table.PreferredWidth = PreferredWidth.FromPercent(50);
            builder.Writeln("Cell #1");

            builder.InsertCell();
            builder.Writeln("Cell #2");

            builder.InsertCell();
            builder.Writeln("Cell #3");

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertTableWithPreferredWidth.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertTableWithPreferredWidth.docx");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.AreEqual(PreferredWidthType.Percent, table.PreferredWidth.Type);
            Assert.AreEqual(50, table.PreferredWidth.Value);
        }

        [Test]
        public void InsertCellsWithPreferredWidths()
        {
            //ExStart
            //ExFor:CellFormat.PreferredWidth
            //ExFor:PreferredWidth
            //ExFor:PreferredWidth.Auto
            //ExFor:PreferredWidth.Equals(PreferredWidth)
            //ExFor:PreferredWidth.Equals(System.Object)
            //ExFor:PreferredWidth.FromPoints
            //ExFor:PreferredWidth.FromPercent
            //ExFor:PreferredWidth.GetHashCode
            //ExFor:PreferredWidth.ToString
            //ExSummary:Shows how to set the different preferred width settings.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table row made up of three cells which have different preferred widths
            Table table = builder.StartTable();

            // Insert an absolute sized cell
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(40);
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightYellow;
            builder.Writeln("Cell at 40 points width");

            PreferredWidth width = builder.CellFormat.PreferredWidth;
            Console.WriteLine($"Width \"{width.GetHashCode()}\": {width.ToString()}");

            // Insert a relative (percent) sized cell
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(20);
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            builder.Writeln("Cell at 20% width");

            // Each cell had its own PreferredWidth
            Assert.False(builder.CellFormat.PreferredWidth.Equals(width));

            width = builder.CellFormat.PreferredWidth;
            Console.WriteLine($"Width \"{width.GetHashCode()}\": {width.ToString()}");

            // Insert a auto sized cell
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
            builder.Writeln(
                "Cell automatically sized. The size of this cell is calculated from the table preferred width.");
            builder.Writeln("In this case the cell will fill up the rest of the available space.");

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertCellsWithPreferredWidths.docx");
            //ExEnd

            Assert.AreEqual(100.0d, PreferredWidth.FromPercent(100).Value);
            Assert.AreEqual(100.0d, PreferredWidth.FromPoints(100).Value);

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertCellsWithPreferredWidths.docx");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);
            
            Assert.AreEqual(PreferredWidthType.Points, table.FirstRow.Cells[0].CellFormat.PreferredWidth.Type);
            Assert.AreEqual(40.0d, table.FirstRow.Cells[0].CellFormat.PreferredWidth.Value);

            Assert.AreEqual(PreferredWidthType.Percent, table.FirstRow.Cells[1].CellFormat.PreferredWidth.Type);
            Assert.AreEqual(20.0d, table.FirstRow.Cells[1].CellFormat.PreferredWidth.Value);

            Assert.AreEqual(PreferredWidthType.Auto, table.FirstRow.Cells[2].CellFormat.PreferredWidth.Type);
            Assert.AreEqual(0.0d, table.FirstRow.Cells[2].CellFormat.PreferredWidth.Value);
        }

        [Test]
        public void InsertTableFromHtml()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the table from HTML. Note that AutoFitSettings does not apply to tables
            // inserted from HTML.
            builder.InsertHtml("<table>" + "<tr>" + "<td>Row 1, Cell 1</td>" + "<td>Row 1, Cell 2</td>" + "</tr>" +
                               "<tr>" + "<td>Row 2, Cell 2</td>" + "<td>Row 2, Cell 2</td>" + "</tr>" + "</table>");

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertTableFromHtml.docx");

            // Verify the table was constructed properly
            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertTableFromHtml.docx");

            Assert.AreEqual(1, doc.GetChildNodes(NodeType.Table, true).Count);
            Assert.AreEqual(2, doc.GetChildNodes(NodeType.Row, true).Count);
            Assert.AreEqual(4, doc.GetChildNodes(NodeType.Cell, true).Count);
        }

        [Test]
        public void InsertNestedTable()
        {
            //ExStart
            //ExFor:Cell.FirstParagraph
            //ExSummary:Shows how to insert a nested table using DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build the outer table
            Cell cell = builder.InsertCell();
            builder.Writeln("Outer Table Cell 1");

            builder.InsertCell();
            builder.Writeln("Outer Table Cell 2");

            // This call is important in order to create a nested table within the first table
            // Without this call the cells inserted below will be appended to the outer table
            builder.EndTable();

            // Move to the first cell of the outer table
            builder.MoveTo(cell.FirstParagraph);

            // Build the inner table
            builder.InsertCell();
            builder.Writeln("Inner Table Cell 1");
            builder.InsertCell();
            builder.Writeln("Inner Table Cell 2");

            builder.EndTable();

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertNestedTable.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertNestedTable.docx");

            Assert.AreEqual(2, doc.GetChildNodes(NodeType.Table, true).Count);
            Assert.AreEqual(4, doc.GetChildNodes(NodeType.Cell, true).Count);
            Assert.AreEqual(1, cell.Tables[0].Count);
            Assert.AreEqual(2, cell.Tables[0].FirstRow.Cells.Count);
        }

        [Test]
        public void CreateSimpleTable()
        {
            //ExStart
            //ExFor:DocumentBuilder
            //ExFor:DocumentBuilder.Write
            //ExFor:DocumentBuilder.InsertCell
            //ExSummary:Shows how to create a simple table using DocumentBuilder with default formatting.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // We call this method to start building the table
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Row 1, Cell 1 Content.");

            // Build the second cell
            builder.InsertCell();
            builder.Write("Row 1, Cell 2 Content.");
            // Call the following method to end the row and start a new row
            builder.EndRow();

            // Build the first cell of the second row
            builder.InsertCell();
            builder.Write("Row 2, Cell 1 Content.");

            // Build the second cell.
            builder.InsertCell();
            builder.Write("Row 2, Cell 2 Content.");
            builder.EndRow();

            // Signal that we have finished building the table
            builder.EndTable();

            // Save the document to disk
            doc.Save(ArtifactsDir + "DocumentBuilder.CreateSimpleTable.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.CreateSimpleTable.docx");
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            Assert.AreEqual(4, table.GetChildNodes(NodeType.Cell, true).Count);

            Assert.AreEqual("Row 1, Cell 1 Content.\a", table.Rows[0].Cells[0].GetText().Trim());
            Assert.AreEqual("Row 1, Cell 2 Content.\a", table.Rows[0].Cells[1].GetText().Trim());
            Assert.AreEqual("Row 2, Cell 1 Content.\a", table.Rows[1].Cells[0].GetText().Trim());
            Assert.AreEqual("Row 2, Cell 2 Content.\a", table.Rows[1].Cells[1].GetText().Trim());

        }

        [Test]
        public void BuildFormattedTable()
        {
            //ExStart
            //ExFor:RowFormat.Height
            //ExFor:RowFormat.HeightRule
            //ExFor:Table.LeftIndent
            //ExFor:DocumentBuilder.ParagraphFormat
            //ExFor:DocumentBuilder.Font
            //ExSummary:Shows how to create a formatted table using DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();

            // Make the header row
            builder.InsertCell();

            // Set the left indent for the table. Table wide formatting must be applied after 
            // at least one row is present in the table
            table.LeftIndent = 20.0;

            // Set height and define the height rule for the header row
            builder.RowFormat.Height = 40.0;
            builder.RowFormat.HeightRule = HeightRule.AtLeast;

            // Some special features for the header row
            builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Font.Size = 16;
            builder.Font.Name = "Arial";
            builder.Font.Bold = true;

            builder.CellFormat.Width = 100.0;
            builder.Write("Header Row,\n Cell 1");

            // We don't need to specify the width of this cell because it's inherited from the previous cell
            builder.InsertCell();
            builder.Write("Header Row,\n Cell 2");

            builder.InsertCell();
            builder.CellFormat.Width = 200.0;
            builder.Write("Header Row,\n Cell 3");
            builder.EndRow();

            // Set features for the other rows and cells
            builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
            builder.CellFormat.Width = 100.0;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

            // Reset height and define a different height rule for table body
            builder.RowFormat.Height = 30.0;
            builder.RowFormat.HeightRule = HeightRule.Auto;
            builder.InsertCell();
            // Reset font formatting
            builder.Font.Size = 12;
            builder.Font.Bold = false;

            // Build the other cells
            builder.Write("Row 1, Cell 1 Content");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2 Content");

            builder.InsertCell();
            builder.CellFormat.Width = 200.0;
            builder.Write("Row 1, Cell 3 Content");
            builder.EndRow();

            builder.InsertCell();
            builder.CellFormat.Width = 100.0;
            builder.Write("Row 2, Cell 1 Content");

            builder.InsertCell();
            builder.Write("Row 2, Cell 2 Content");

            builder.InsertCell();
            builder.CellFormat.Width = 200.0;
            builder.Write("Row 2, Cell 3 Content.");
            builder.EndRow();
            builder.EndTable();

            doc.Save(ArtifactsDir + "DocumentBuilder.CreateFormattedTable.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.CreateFormattedTable.docx");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.AreEqual(20.0d, table.LeftIndent);

            Assert.AreEqual(HeightRule.AtLeast, table.Rows[0].RowFormat.HeightRule);
            Assert.AreEqual(40.0d, table.Rows[0].RowFormat.Height);

            foreach (Cell c in doc.GetChildNodes(NodeType.Cell, true))
            {
                Assert.AreEqual(ParagraphAlignment.Center, c.FirstParagraph.ParagraphFormat.Alignment);

                foreach (Run r in c.FirstParagraph.Runs)
                {
                    Assert.AreEqual("Arial", r.Font.Name);

                    if (c.ParentRow == table.FirstRow)
                    {
                        Assert.AreEqual(16, r.Font.Size);
                        Assert.True(r.Font.Bold);
                    }
                    else
                    {
                        Assert.AreEqual(12, r.Font.Size);
                        Assert.False(r.Font.Bold);
                    }
                }
            }
        }

        [Test]
        public void TableBordersAndShading()
        {
            //ExStart
            //ExFor:Shading
            //ExFor:Table.SetBorders
            //ExFor:BorderCollection.Left
            //ExFor:BorderCollection.Right
            //ExFor:BorderCollection.Top
            //ExFor:BorderCollection.Bottom
            //ExSummary:Shows how to format table and cell with different borders and shadings.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table and set a default color/thickness for its borders
            Table table = builder.StartTable();
            table.SetBorders(LineStyle.Single, 2.0, Color.Black);

            // Set the cell shading for this cell
            builder.InsertCell();
            builder.CellFormat.Shading.BackgroundPatternColor = Color.Red;
            builder.Writeln("Cell #1");

            // Specify a different cell shading for the second cell
            builder.InsertCell();
            builder.CellFormat.Shading.BackgroundPatternColor = Color.Green;
            builder.Writeln("Cell #2");

            // End this row
            builder.EndRow();

            // Clear the cell formatting from previous operations
            builder.CellFormat.ClearFormatting();

            // Create the second row
            builder.InsertCell();
            builder.Writeln("Cell #3");

            // Clear the cell formatting from the previous cell
            builder.CellFormat.ClearFormatting();

            builder.CellFormat.Borders.Left.LineWidth = 4.0;
            builder.CellFormat.Borders.Right.LineWidth = 4.0;
            builder.CellFormat.Borders.Top.LineWidth = 4.0;
            builder.CellFormat.Borders.Bottom.LineWidth = 4.0;

            builder.InsertCell();
            builder.Writeln("Cell #4");

            doc.Save(ArtifactsDir + "DocumentBuilder.TableBordersAndShading.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.TableBordersAndShading.docx");
            table = (Table) doc.GetChild(NodeType.Table, 0, true);

            foreach (Cell c in table.FirstRow)
            {
                Assert.AreEqual(0.5d, c.CellFormat.Borders.Top.LineWidth);
                Assert.AreEqual(0.5d, c.CellFormat.Borders.Bottom.LineWidth);
                Assert.AreEqual(0.5d, c.CellFormat.Borders.Left.LineWidth);
                Assert.AreEqual(0.5d, c.CellFormat.Borders.Right.LineWidth);

                Assert.AreEqual(Color.Empty.ToArgb(), c.CellFormat.Borders.Left.Color.ToArgb());
                Assert.AreEqual(LineStyle.Single, c.CellFormat.Borders.Left.LineStyle);
            }

            Assert.AreEqual(Color.Red.ToArgb(),
                table.FirstRow.FirstCell.CellFormat.Shading.BackgroundPatternColor.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(),
                table.FirstRow.Cells[1].CellFormat.Shading.BackgroundPatternColor.ToArgb());

            foreach (Cell c in table.LastRow)
            {
                Assert.AreEqual(4.0d, c.CellFormat.Borders.Top.LineWidth);
                Assert.AreEqual(4.0d, c.CellFormat.Borders.Bottom.LineWidth);
                Assert.AreEqual(4.0d, c.CellFormat.Borders.Left.LineWidth);
                Assert.AreEqual(4.0d, c.CellFormat.Borders.Right.LineWidth);

                Assert.AreEqual(Color.Empty.ToArgb(), c.CellFormat.Borders.Left.Color.ToArgb());
                Assert.AreEqual(LineStyle.Single, c.CellFormat.Borders.Left.LineStyle);
                Assert.AreEqual(Color.Empty.ToArgb(), c.CellFormat.Shading.BackgroundPatternColor.ToArgb());
            }
        }

        [Test]
        public void SetPreferredTypeConvertUtil()
        {
            //ExStart
            //ExFor:PreferredWidth.FromPoints
            //ExSummary:Shows how to specify a cell preferred width by converting inches to points.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(ConvertUtil.InchToPoint(3));
            builder.InsertCell();
            //ExEnd

            Assert.AreEqual(216.0, table.FirstRow.FirstCell.CellFormat.PreferredWidth.Value);
        }

        [Test]
        public void InsertHyperlinkToLocalBookmark()
        {
            //ExStart
            //ExFor:DocumentBuilder.StartBookmark
            //ExFor:DocumentBuilder.EndBookmark
            //ExFor:DocumentBuilder.InsertHyperlink
            //ExSummary:Shows how to insert a hyperlink referencing a local bookmark.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartBookmark("Bookmark1");
            builder.Write("Bookmarked text.");
            builder.EndBookmark("Bookmark1");

            builder.Writeln("Some other text");

            // Specify font formatting for the hyperlink
            builder.Font.Color = Color.Blue;
            builder.Font.Underline = Underline.Single;

            // Insert hyperlink
            // Switch \o is used to provide hyperlink tip text
            builder.InsertHyperlink("Hyperlink Text", @"Bookmark1"" \o ""Hyperlink Tip", true);

            // Clear hyperlink formatting
            builder.Font.ClearFormatting();

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertHyperlinkToLocalBookmark.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertHyperlinkToLocalBookmark.docx");
            FieldHyperlink hyperlink = (FieldHyperlink)doc.Range.Fields[0];

            TestUtil.VerifyField(FieldType.FieldHyperlink, " HYPERLINK \\l \"Bookmark1\" \\o \"Hyperlink Tip\" ", "Hyperlink Text", hyperlink);
            Assert.AreEqual("Bookmark1", hyperlink.SubAddress);
            Assert.IsTrue(doc.Range.Bookmarks.Any(b => b.Name == "Bookmark1"));
        }

        [Test]
        public void DocumentBuilderCursorPosition()
        {
            // Write some text in a blank Document using a DocumentBuilder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Hello world!");

            // If the builder's cursor is at the end of the document, there will be no nodes in front of it so the current node will be null
            Assert.Null(builder.CurrentNode);

            // However, the current paragraph the cursor is in will be valid
            Assert.AreEqual("Hello world!", builder.CurrentParagraph.GetText().Trim());

            // Move to the beginning of the document and place the cursor at an existing node
            builder.MoveToDocumentStart();          
            Assert.AreEqual(NodeType.Run, builder.CurrentNode.NodeType);
        }

        [Test]
        public void DocumentBuilderMoveToNode()
        {
            //ExStart
            //ExFor:Story.LastParagraph
            //ExFor:DocumentBuilder.MoveTo(Node)
            //ExSummary:Shows how to move a DocumentBuilder's cursor position to a specified node.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            // Write a paragraph with the DocumentBuilder
            builder.Writeln("Text 1. ");

            // Move the DocumentBuilder to the first paragraph of the document and add another paragraph
            Assert.AreEqual(doc.FirstSection.Body.LastParagraph, builder.CurrentParagraph); //ExSkip
            builder.MoveTo(doc.FirstSection.Body.FirstParagraph.Runs[0]);
            Assert.AreEqual(doc.FirstSection.Body.FirstParagraph, builder.CurrentParagraph); //ExSkip
            builder.Writeln("Text 2. ");

            // Since we moved to a node before the first paragraph before we added a second paragraph,
            // the second paragraph will appear in front of the first paragraph
            Assert.AreEqual("Text 2. \rText 1.", doc.GetText().Trim());

            // We can move the DocumentBuilder back to the end of the document like this
            // and carry on adding text to the end of the document
            builder.MoveTo(doc.FirstSection.Body.LastParagraph);
            builder.Writeln("Text 3. ");

            Assert.AreEqual("Text 2. \rText 1. \rText 3.", doc.GetText().Trim());
            Assert.AreEqual(doc.FirstSection.Body.LastParagraph, builder.CurrentParagraph); //ExSkip
            //ExEnd
        }

        [Test]
        public void DocumentBuilderMoveToDocumentStartEnd()
        {
            Document doc = new Document(MyDir + "Document.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToDocumentEnd();
            builder.Writeln("This is the end of the document.");

            builder.MoveToDocumentStart();
            builder.Writeln("This is the beginning of the document.");
        }

        [Test]
        public void DocumentBuilderMoveToSection()
        {
            // Create a blank document and append a section to it, giving it two sections
            Document doc = new Document();
            doc.AppendChild(new Section(doc));

            // Move a DocumentBuilder to the second section and add text
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToSection(1);
            builder.Writeln("Text added to the 2nd section.");
        }

        [Test]
        public void DocumentBuilderMoveToParagraph()
        {
            //ExStart
            //ExFor:DocumentBuilder.MoveToParagraph
            //ExSummary:Shows how to move a cursor position to the specified paragraph.
            // Open a document with a lot of paragraphs
            Document doc = new Document(MyDir + "Paragraphs.docx");
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;

            Assert.AreEqual(22, paragraphs.Count);

            // When we create a DocumentBuilder for a document, its cursor is at the very beginning of the document by default,
            // and any content added by the DocumentBuilder will just be prepended to the document
            DocumentBuilder builder = new DocumentBuilder(doc);

            Assert.AreEqual(0, paragraphs.IndexOf(builder.CurrentParagraph));

            // We can manually move the DocumentBuilder to any paragraph in the document via a 0-based index like this
            builder.MoveToParagraph(2, 0);
            Assert.AreEqual(2, paragraphs.IndexOf(builder.CurrentParagraph)); //ExSkip
            builder.Writeln("This is a new third paragraph. ");
            //ExEnd

            Assert.AreEqual(3, paragraphs.IndexOf(builder.CurrentParagraph));

            doc = DocumentHelper.SaveOpen(doc);

            Assert.AreEqual("This is a new third paragraph.", doc.FirstSection.Body.Paragraphs[2].GetText().Trim());
        }

        [Test]
        public void DocumentBuilderMoveToTableCell()
        {
            //ExStart
            //ExFor:DocumentBuilder.MoveToCell
            //ExSummary:Shows how to move a cursor position to the specified table cell.
            Document doc = new Document(MyDir + "Tables.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the builder to row 3, cell 4 of the first table
            builder.MoveToCell(0, 2, 3, 0);
            builder.Write("\nCell contents added by DocumentBuilder");
            //ExEnd

            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.AreEqual(table.Rows[2].Cells[3], builder.CurrentNode.ParentNode.ParentNode);
            Assert.AreEqual("Cell contents added by DocumentBuilderCell 3 contents\a", table.Rows[2].Cells[3].GetText().Trim());

        }

        [Test]
        public void DocumentBuilderMoveToBookmarkEnd()
        {
            //ExStart
            //ExFor:DocumentBuilder.MoveToBookmark(String, Boolean, Boolean)
            //ExSummary:Shows how to move a cursor position to just after the bookmark end.
            Document doc = new Document(MyDir + "Bookmarks.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move to after the end of the first bookmark
            Assert.True(builder.MoveToBookmark("MyBookmark1", false, true));
            builder.Write(" Text appended via DocumentBuilder.");
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);

            Assert.False(doc.Range.Bookmarks["MyBookmark1"].Text.Contains(" Text appended via DocumentBuilder."));
        }

        [Test]
        public void DocumentBuilderBuildTable()
        {
            //ExStart
            //ExFor:Table
            //ExFor:DocumentBuilder.StartTable
            //ExFor:DocumentBuilder.EndRow
            //ExFor:DocumentBuilder.EndTable
            //ExFor:DocumentBuilder.CellFormat
            //ExFor:DocumentBuilder.RowFormat
            //ExFor:DocumentBuilder.Write(String)
            //ExFor:DocumentBuilder.Writeln(String)
            //ExFor:CellVerticalAlignment
            //ExFor:CellFormat.Orientation
            //ExFor:TextOrientation
            //ExFor:Table.AutoFit
            //ExFor:AutoFitBehavior
            //ExSummary:Shows how to build a formatted table that contains 2 rows and 2 columns.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();

            // Insert a cell
            builder.InsertCell();
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.Write("This is row 1 cell 1");

            // Use fixed column widths
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // Insert a cell
            builder.InsertCell();
            builder.Write("This is row 1 cell 2");
            builder.EndRow();

            // Insert a cell
            builder.InsertCell();

            // Apply new row formatting
            builder.RowFormat.Height = 100;
            builder.RowFormat.HeightRule = HeightRule.Exactly;

            builder.CellFormat.Orientation = TextOrientation.Upward;
            builder.Write("This is row 2 cell 1");

            // Insert a cell
            builder.InsertCell();
            builder.CellFormat.Orientation = TextOrientation.Downward;
            builder.Write("This is row 2 cell 2");

            builder.EndRow();
            builder.EndTable();
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.AreEqual(2, table.Rows.Count);
            Assert.AreEqual(2, table.Rows[0].Cells.Count);
            Assert.AreEqual(2, table.Rows[1].Cells.Count);
            Assert.False(table.AllowAutoFit);

            Assert.AreEqual(0, table.Rows[0].RowFormat.Height);
            Assert.AreEqual(HeightRule.Auto, table.Rows[0].RowFormat.HeightRule);
            Assert.AreEqual(100, table.Rows[1].RowFormat.Height);
            Assert.AreEqual(HeightRule.Exactly, table.Rows[1].RowFormat.HeightRule);

            Assert.AreEqual("This is row 1 cell 1\a", table.Rows[0].Cells[0].GetText().Trim());
            Assert.AreEqual(CellVerticalAlignment.Center, table.Rows[0].Cells[0].CellFormat.VerticalAlignment);

            Assert.AreEqual("This is row 1 cell 2\a", table.Rows[0].Cells[1].GetText().Trim());

            Assert.AreEqual("This is row 2 cell 1\a", table.Rows[1].Cells[0].GetText().Trim());
            Assert.AreEqual(TextOrientation.Upward, table.Rows[1].Cells[0].CellFormat.Orientation);

            Assert.AreEqual("This is row 2 cell 2\a", table.Rows[1].Cells[1].GetText().Trim());
            Assert.AreEqual(TextOrientation.Downward, table.Rows[1].Cells[1].CellFormat.Orientation);
        }

        [Test]
        public void TableCellVerticalRotatedFarEastTextOrientation()
        {
            Document doc = new Document(MyDir + "Rotated cell text.docx");

            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            Cell cell = table.FirstRow.FirstCell;

            Assert.AreEqual(TextOrientation.VerticalRotatedFarEast, cell.CellFormat.Orientation);

            doc = DocumentHelper.SaveOpen(doc);

            table = (Table) doc.GetChild(NodeType.Table, 0, true);
            cell = table.FirstRow.FirstCell;

            Assert.AreEqual(TextOrientation.VerticalRotatedFarEast, cell.CellFormat.Orientation);
        }

        [Test]
        public void DocumentBuilderInsertBreak()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("This is page 1.");
            builder.InsertBreak(BreakType.PageBreak);

            builder.Writeln("This is page 2.");
            builder.InsertBreak(BreakType.PageBreak);

            builder.Writeln("This is page 3.");
        }

        [Test]
        public void DocumentBuilderInsertInlineImage()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage(ImageDir + "Transparent background logo.png");
        }

        [Test]
        public void DocumentBuilderInsertFloatingImage()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows how to insert a floating image from a file or URL.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage(ImageDir + "Transparent background logo.png", RelativeHorizontalPosition.Margin, 100,
                RelativeVerticalPosition.Margin, 100, 200, 100, WrapType.Square);
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            Shape image = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Png, image);
            Assert.AreEqual(100.0d, image.Left);
            Assert.AreEqual(100.0d, image.Top);
            Assert.AreEqual(200.0d, image.Width);
            Assert.AreEqual(100.0d, image.Height);
            Assert.AreEqual(WrapType.Square, image.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Margin, image.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Margin, image.RelativeVerticalPosition);
        }

        [Test]
        public void InsertImageFromUrl()
        {
            // Insert an image from a URL
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(AsposeLogoUrl);

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertImageFromUrl.doc");

            // Verify that the image was inserted into the document
            Shape shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            Assert.IsNotNull(shape);
            Assert.True(shape.HasImage);
        }

        [Test]
        public void InsertImageOriginalSize()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows how to insert a floating image from a file or URL and retain the original image size in the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Pass a negative value to the width and height values to specify using the size of the source image
            builder.InsertImage(ImageDir + "Logo.jpg", RelativeHorizontalPosition.Margin, 200,
                RelativeVerticalPosition.Margin, 100, -1, -1, WrapType.Square);
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            Shape image = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(400, 400, ImageType.Jpeg, image);
            Assert.AreEqual(200.0d, image.Left);
            Assert.AreEqual(100.0d, image.Top);
            Assert.AreEqual(270.3d, image.Width);
            Assert.AreEqual(270.3d, image.Height);
            Assert.AreEqual(WrapType.Square, image.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Margin, image.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Margin, image.RelativeVerticalPosition);
        }

        [Test]
        public void DocumentBuilderInsertTextInputFormField()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertTextInput
            //ExSummary:Shows how to insert a text input form field into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Hello", 0);
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            FormField formField = doc.Range.FormFields[0];

            Assert.True(formField.Enabled);
            Assert.AreEqual("TextInput", formField.Name);
            Assert.AreEqual(0, formField.MaxLength);
            Assert.AreEqual("Hello", formField.Result);
            Assert.AreEqual(FieldType.FieldFormTextInput, formField.Type);
            Assert.AreEqual("", formField.TextInputFormat);
            Assert.AreEqual(TextFormFieldType.Regular, formField.TextInputType);
        }

        [Test]
        public void DocumentBuilderInsertComboBoxFormField()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertComboBox
            //ExSummary:Shows how to insert a combobox form field into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            string[] items = { "One", "Two", "Three" };
            builder.InsertComboBox("DropDown", items, 0);
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            FormField formField = doc.Range.FormFields[0];

            Assert.True(formField.Enabled);
            Assert.AreEqual("DropDown", formField.Name);
            Assert.AreEqual(0, formField.DropDownSelectedIndex);
            Assert.AreEqual(new[] { "One", "Two", "Three" } , formField.DropDownItems);
            Assert.AreEqual(FieldType.FieldFormDropDown, formField.Type);
        }

        [Test]
        public void DocumentBuilderInsertToc()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table of contents at the beginning of the document
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

            // The newly inserted table of contents will be initially empty
            // It needs to be populated by updating the fields in the document
            doc.UpdateFields();
        }

        [Test]
        [Description("WORDSNET-16868")]
        public void SignatureLineProviderId()
        {
            //ExStart
            //ExFor:SignatureLine.ProviderId
            //ExFor:SignatureLineOptions.ShowDate
            //ExFor:SignatureLineOptions.Email
            //ExFor:SignatureLineOptions.DefaultInstructions
            //ExFor:SignatureLineOptions.Instructions
            //ExFor:SignatureLineOptions.AllowComments
            //ExFor:DocumentBuilder.InsertSignatureLine(SignatureLineOptions)
            //ExFor:SignOptions.ProviderId
            //ExSummary:Shows how to sign document with personal certificate and specific signature line.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            SignatureLineOptions signatureLineOptions = new SignatureLineOptions
            {
                Signer = "vderyushev",
                SignerTitle = "QA",
                Email = "vderyushev@aspose.com",
                ShowDate = true,
                DefaultInstructions = false,
                Instructions = "You need more info about signature line",
                AllowComments = true
            };

            SignatureLine signatureLine = builder.InsertSignatureLine(signatureLineOptions).SignatureLine;
            signatureLine.ProviderId = Guid.Parse("CF5A7BB4-8F3C-4756-9DF6-BEF7F13259A2");
            
            doc.Save(ArtifactsDir + "DocumentBuilder.SignatureLineProviderId.docx");

            SignOptions signOptions = new SignOptions
            {
                SignatureLineId = signatureLine.Id,
                ProviderId = signatureLine.ProviderId,
                Comments = "Document was signed by vderyushev",
                SignTime = DateTime.Now
            };

            CertificateHolder certHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");

            DigitalSignatureUtil.Sign(ArtifactsDir + "DocumentBuilder.SignatureLineProviderId.docx", 
                ArtifactsDir + "DocumentBuilder.SignatureLineProviderId.Signed.docx", certHolder, signOptions);
            //ExEnd
            
            doc = new Document(ArtifactsDir + "DocumentBuilder.SignatureLineProviderId.Signed.docx");
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            signatureLine = shape.SignatureLine;

            Assert.AreEqual("vderyushev", signatureLine.Signer);
            Assert.AreEqual("QA", signatureLine.SignerTitle);
            Assert.AreEqual("vderyushev@aspose.com", signatureLine.Email);
            Assert.True(signatureLine.ShowDate);
            Assert.False(signatureLine.DefaultInstructions);
            Assert.AreEqual("You need more info about signature line", signatureLine.Instructions);
            Assert.True(signatureLine.AllowComments);
            Assert.True(signatureLine.IsSigned);
            Assert.True(signatureLine.IsValid);

            DigitalSignatureCollection signatures = DigitalSignatureUtil.LoadSignatures(
                ArtifactsDir + "DocumentBuilder.SignatureLineProviderId.Signed.docx");

            Assert.AreEqual(1, signatures.Count);
            Assert.True(signatures[0].IsValid);
            Assert.AreEqual("Document was signed by vderyushev", signatures[0].Comments);
            Assert.AreEqual(DateTime.Today, signatures[0].SignTime.Date);
            Assert.AreEqual("CN=Morzal.Me", signatures[0].IssuerName);
            Assert.AreEqual(DigitalSignatureType.XmlDsig, signatures[0].SignatureType);

        }

        [Test]
        public void InsertSignatureLineCurrentPosition()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertSignatureLine(SignatureLineOptions, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, WrapType)
            //ExSummary:Shows how to insert signature line at the specified position.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            SignatureLineOptions options = new SignatureLineOptions
            {
                Signer = "John Doe",
                SignerTitle = "Manager",
                Email = "johndoe@aspose.com",
                ShowDate = true,
                DefaultInstructions = false,
                Instructions = "You need more info about signature line",
                AllowComments = true
            };

            builder.InsertSignatureLine(options, RelativeHorizontalPosition.RightMargin, 2.0,
                RelativeVerticalPosition.Page, 3.0, WrapType.Inline);
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);

            Shape shape = (Shape) doc.GetChild(NodeType.Shape, 0, true);
            SignatureLine signatureLine = shape.SignatureLine;

            Assert.AreEqual("John Doe", signatureLine.Signer);
            Assert.AreEqual("Manager", signatureLine.SignerTitle);
            Assert.AreEqual("johndoe@aspose.com", signatureLine.Email);
            Assert.True(signatureLine.ShowDate);
            Assert.False(signatureLine.DefaultInstructions);
            Assert.AreEqual("You need more info about signature line", signatureLine.Instructions);
            Assert.True(signatureLine.AllowComments);
            Assert.False(signatureLine.IsSigned);
            Assert.False(signatureLine.IsValid);
        }

        [Test]
        public void DocumentBuilderSetFontFormatting()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set font formatting properties
            Aspose.Words.Font font = builder.Font;
            font.Bold = true;
            font.Color = Color.DarkBlue;
            font.Italic = true;
            font.Name = "Arial";
            font.Size = 24;
            font.Spacing = 5;
            font.Underline = Underline.Double;

            // Output formatted text
            builder.Writeln("I'm a very nice formatted String.");
        }

        [Test]
        public void DocumentBuilderSetParagraphFormatting()
        {
            //ExStart
            //ExFor:ParagraphFormat.RightIndent
            //ExFor:ParagraphFormat.LeftIndent
            //ExSummary:Shows how to set paragraph formatting.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set paragraph formatting properties
            ParagraphFormat paragraphFormat = builder.ParagraphFormat;
            paragraphFormat.Alignment = ParagraphAlignment.Center;
            paragraphFormat.LeftIndent = 50;
            paragraphFormat.RightIndent = 50;
            paragraphFormat.SpaceAfter = 25;

            // Output text
            builder.Writeln(
                "This paragraph demonstrates how the left and right indents affect word wrapping.");
            builder.Writeln(
                "The space between the above paragraph and this one depends on the DocumentBuilder's paragraph format.");

            doc.Save(ArtifactsDir + "DocumentBuilder.DocumentBuilderSetParagraphFormatting.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.DocumentBuilderSetParagraphFormatting.docx");

            foreach (Paragraph paragraph in doc.FirstSection.Body.Paragraphs)
            {
                Assert.AreEqual(ParagraphAlignment.Center, paragraph.ParagraphFormat.Alignment);
                Assert.AreEqual(50.0d, paragraph.ParagraphFormat.LeftIndent);
                Assert.AreEqual(50.0d, paragraph.ParagraphFormat.RightIndent);
                Assert.AreEqual(25.0d, paragraph.ParagraphFormat.SpaceAfter);

            }
        }

        [Test]
        public void DocumentBuilderSetCellFormatting()
        {
            //ExStart
            //ExFor:DocumentBuilder.CellFormat
            //ExFor:CellFormat.Width
            //ExFor:CellFormat.LeftPadding
            //ExFor:CellFormat.RightPadding
            //ExFor:CellFormat.TopPadding
            //ExFor:CellFormat.BottomPadding
            //ExFor:DocumentBuilder.StartTable
            //ExFor:DocumentBuilder.EndTable
            //ExSummary:Shows how to create a table that contains a single formatted cell.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartTable();
            builder.InsertCell();

            // Set the cell formatting
            CellFormat cellFormat = builder.CellFormat;
            cellFormat.Width = 250;
            cellFormat.LeftPadding = 30;
            cellFormat.RightPadding = 30;
            cellFormat.TopPadding = 30;
            cellFormat.BottomPadding = 30;

            builder.Write("Formatted cell");
            builder.EndRow();
            builder.EndTable();

            doc.Save(ArtifactsDir + "DocumentBuilder.DocumentBuilderSetCellFormatting.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.DocumentBuilderSetCellFormatting.docx");
            Cell firstCell = ((Table)doc.GetChild(NodeType.Table,0, true)).FirstRow.FirstCell;

            Assert.AreEqual("Formatted cell\a", firstCell.GetText().Trim());

            Assert.AreEqual(250.0d, firstCell.CellFormat.Width);
            Assert.AreEqual(30.0d, firstCell.CellFormat.LeftPadding);
            Assert.AreEqual(30.0d, firstCell.CellFormat.RightPadding);
            Assert.AreEqual(30.0d, firstCell.CellFormat.TopPadding);
            Assert.AreEqual(30.0d, firstCell.CellFormat.BottomPadding);

        }

        [Test]
        public void DocumentBuilderSetRowFormatting()
        {
            //ExStart
            //ExFor:DocumentBuilder.RowFormat
            //ExFor:HeightRule
            //ExFor:RowFormat.Height
            //ExFor:RowFormat.HeightRule
            //ExFor:Table.LeftPadding
            //ExFor:Table.RightPadding
            //ExFor:Table.TopPadding
            //ExFor:Table.BottomPadding
            //ExSummary:Shows how to create a table that contains a single cell and apply row formatting.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();

            // Set the row formatting
            RowFormat rowFormat = builder.RowFormat;
            rowFormat.Height = 100;
            rowFormat.HeightRule = HeightRule.Exactly;
            // These formatting properties are set on the table and are applied to all rows in the table
            table.LeftPadding = 30;
            table.RightPadding = 30;
            table.TopPadding = 30;
            table.BottomPadding = 30;

            builder.Writeln("Contents of formatted row.");

            builder.EndRow();
            builder.EndTable();

            doc.Save(ArtifactsDir + "DocumentBuilder.DocumentBuilderSetRowFormatting.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.DocumentBuilderSetRowFormatting.docx");
            table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.AreEqual(30.0d, table.LeftPadding);
            Assert.AreEqual(30.0d, table.RightPadding);
            Assert.AreEqual(30.0d, table.TopPadding);
            Assert.AreEqual(30.0d, table.BottomPadding);

            Assert.AreEqual(100.0d, table.FirstRow.RowFormat.Height);
            Assert.AreEqual(HeightRule.Exactly, table.FirstRow.RowFormat.HeightRule);
        }

        [Test]
        public void DocumentBuilderSetListFormatting()
        {
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
        }

        [Test]
        public void DocumentBuilderSetSectionFormatting()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set page properties
            builder.PageSetup.Orientation = Orientation.Landscape;
            builder.PageSetup.LeftMargin = 50;
            builder.PageSetup.PaperSize = PaperSize.Paper10x14;
        }

        [Test]
        public void InsertFootnote()
        {
            //ExStart
            //ExFor:FootnoteType
            //ExFor:DocumentBuilder.InsertFootnote(FootnoteType,String)
            //ExFor:DocumentBuilder.InsertFootnote(FootnoteType,String,String)
            //ExSummary:Shows how to reference text with a footnote and an endnote.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            // Insert some text and mark it with a footnote with the IsAuto attribute set to "true" by default,
            // so the marker seen in the body text will be auto-numbered at "1", and the footnote will appear at the bottom of the page
            builder.Write("This text will be referenced by a footnote.");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote comment regarding referenced text.");

            // Insert more text and mark it with an endnote with a custom reference mark,
            // which will be used in place of the number "2" and will set "IsAuto" to false
            builder.Write("This text will be referenced by an endnote.");
            builder.InsertFootnote(FootnoteType.Endnote, "Endnote comment regarding referenced text.", "CustomMark");

            // Footnotes always appear at the bottom of the page of their referenced text, so this page break will not affect the footnote
            // On the other hand, endnotes are always at the end of the document, so this page break will push the endnote down to the next page
            builder.InsertBreak(BreakType.PageBreak);

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertFootnote.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertFootnote.docx");

            TestUtil.VerifyFootnote(FootnoteType.Footnote, true, string.Empty,
                "Footnote comment regarding referenced text.", (Footnote)doc.GetChild(NodeType.Footnote, 0, true));
            TestUtil.VerifyFootnote(FootnoteType.Endnote, false, "CustomMark",
                "CustomMark Endnote comment regarding referenced text.", (Footnote)doc.GetChild(NodeType.Footnote, 1, true));
        }

        [Test]
        public void DocumentBuilderApplyParagraphStyle()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set paragraph style
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;

            builder.Write("Hello");
        }

        [Test]
        public void DocumentBuilderApplyBordersAndShading()
        {
            //ExStart
            //ExFor:BorderCollection.Item(BorderType)
            //ExFor:Shading
            //ExFor:TextureIndex
            //ExFor:ParagraphFormat.Shading
            //ExFor:Shading.Texture
            //ExFor:Shading.BackgroundPatternColor
            //ExFor:Shading.ForegroundPatternColor
            //ExSummary:Shows how to apply borders and shading to a paragraph.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set paragraph borders
            BorderCollection borders = builder.ParagraphFormat.Borders;
            borders.DistanceFromText = 20;
            borders[BorderType.Left].LineStyle = LineStyle.Double;
            borders[BorderType.Right].LineStyle = LineStyle.Double;
            borders[BorderType.Top].LineStyle = LineStyle.Double;
            borders[BorderType.Bottom].LineStyle = LineStyle.Double;

            // Set paragraph shading
            Shading shading = builder.ParagraphFormat.Shading;
            shading.Texture = TextureIndex.TextureDiagonalCross;
            shading.BackgroundPatternColor = Color.LightCoral;
            shading.ForegroundPatternColor = Color.LightSalmon;

            builder.Write("This paragraph is formatted with a double border and shading.");
            doc.Save(ArtifactsDir + "DocumentBuilder.DocumentBuilderApplyBordersAndShading.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.DocumentBuilderApplyBordersAndShading.docx");
            borders = doc.FirstSection.Body.FirstParagraph.ParagraphFormat.Borders;

            Assert.AreEqual(20.0d, borders.DistanceFromText);
            Assert.AreEqual(LineStyle.Double, borders[BorderType.Left].LineStyle);
            Assert.AreEqual(LineStyle.Double, borders[BorderType.Right].LineStyle);
            Assert.AreEqual(LineStyle.Double, borders[BorderType.Top].LineStyle);
            Assert.AreEqual(LineStyle.Double, borders[BorderType.Bottom].LineStyle);

            Assert.AreEqual(TextureIndex.TextureDiagonalCross, shading.Texture);
            Assert.AreEqual(Color.LightCoral.ToArgb(), shading.BackgroundPatternColor.ToArgb());
            Assert.AreEqual(Color.LightSalmon.ToArgb(), shading.ForegroundPatternColor.ToArgb());

        }

        [Test]
        public void DeleteRow()
        {
            //ExStart
            //ExFor:DocumentBuilder.DeleteRow
            //ExSummary:Shows how to delete a row from a table.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table with 2 rows
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndRow();
            builder.InsertCell();
            builder.Write("Cell 3");
            builder.InsertCell();
            builder.Write("Cell 4");
            builder.EndTable();

            Assert.AreEqual(2, table.Rows.Count);

            // Delete the first row of the first table in the document
            builder.DeleteRow(0, 0);

            Assert.AreEqual(1, table.Rows.Count);
            //ExEnd

            Assert.AreEqual("Cell 3\aCell 4\a\a", table.GetText().Trim());
        }

        [Test]
        [Ignore("Bug: does not insert headers and footers, all lists (bullets, numbering, multilevel) breaks")]
        public void InsertDocument()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertDocument(Document, ImportFormatMode)
            //ExFor:ImportFormatMode
            //ExSummary:Shows how to insert a document content into another document keep formatting of inserted document.
            Document doc = new Document(MyDir + "Document.docx");

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();
            builder.InsertBreak(BreakType.PageBreak);

            Document docToInsert = new Document(MyDir + "Formatted elements.docx");

            builder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);
            builder.Document.Save(ArtifactsDir + "DocumentBuilder.InsertDocument.docx");
            //ExEnd

            Assert.AreEqual(29, doc.Styles.Count);
            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "DocumentBuilder.InsertDocument.docx", 
                GoldsDir + "DocumentBuilder.InsertDocument Gold.docx"));
        }

        [Test]
        public void KeepSourceNumbering()
        {
            //ExStart
            //ExFor:ImportFormatOptions.KeepSourceNumbering
            //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode, ImportFormatOptions)
            //ExSummary:Shows how the numbering will be imported when it clashes in source and destination documents.
            // Open a document with a custom list numbering scheme and clone it
            // Since both have the same numbering format, the formats will clash if we import one document into the other
            Document srcDoc = new Document(MyDir + "Custom list numbering.docx");
            Document dstDoc = srcDoc.Clone();
            
            // Both documents have the same numbering in their lists, but if we set this flag to false and then import one document into the other
            // the numbering of the imported source document will continue from where it ends in the destination document
            ImportFormatOptions importFormatOptions = new ImportFormatOptions();
            importFormatOptions.KeepSourceNumbering = false;

            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepDifferentStyles, importFormatOptions);
            foreach (Paragraph paragraph in srcDoc.FirstSection.Body.Paragraphs)
            {
                Node importedNode = importer.ImportNode(paragraph, true);
                dstDoc.FirstSection.Body.AppendChild(importedNode);
            }
            
            dstDoc.UpdateListLabels();
            dstDoc.Save(ArtifactsDir + "DocumentBuilder.KeepSourceNumbering.docx");
            //ExEnd
        }


        [Test]
        public void ResolveStyleBehaviorWhileAppendDocument()
        {
            //ExStart
            //ExFor:Document.AppendDocument(Document, ImportFormatMode, ImportFormatOptions)
            //ExSummary:Shows how to resolve styles behavior while append document.
            // Open a document with text in a custom style and clone it
            Document srcDoc = new Document(MyDir + "Custom list numbering.docx");
            Document dstDoc = srcDoc.Clone();

            // We now have two documents, each with an identical style named "CustomStyle" 
            // We can change the text color of one of the styles
            dstDoc.Styles["CustomStyle"].Font.Color = Color.DarkRed;

            ImportFormatOptions options = new ImportFormatOptions();
            // Specify that if numbering clashes in source and destination documents
            // then a numbering from the source document will be used
            options.KeepSourceNumbering = true;

            // If we join two documents which have different styles that share the same name,
            // we can resolve the style clash with an ImportFormatMode
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepDifferentStyles, options);
            dstDoc.UpdateListLabels();

            dstDoc.Save(ArtifactsDir + "DocumentBuilder.ResolveStyleBehaviorWhileAppendDocument.docx");
            //ExEnd
        }

        [TestCase(true)]
        [TestCase(false)]
        public void IgnoreTextBoxes(bool isIgnoreTextBoxes)
        {
            //ExStart
            //ExFor:ImportFormatOptions.IgnoreTextBoxes
            //ExSummary:Shows how to manage formatting in the text boxes of the source destination during the import.
            // Create a document and add text
            Document dstDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(dstDoc);

            builder.Writeln("Hello world! Text box to follow.");

            // Create another document with a textbox, and insert some formatted text into it
            Document srcDoc = new Document();
            builder = new DocumentBuilder(srcDoc);

            Shape textBox = builder.InsertShape(ShapeType.TextBox, 300, 100);
            builder.MoveTo(textBox.FirstParagraph);
            builder.ParagraphFormat.Style.Font.Name = "Courier New";
            builder.ParagraphFormat.Style.Font.Size = 24.0d;
            builder.Write("Textbox contents");

            // When we import the document with the textbox as a node into the first document, by default the text inside the text box will keep its formatting
            // Setting the IgnoreTextBoxes flag will clear the formatting during importing of the node
            ImportFormatOptions importFormatOptions = new ImportFormatOptions();
            importFormatOptions.IgnoreTextBoxes = isIgnoreTextBoxes;

            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting, importFormatOptions);

            foreach (Paragraph paragraph in srcDoc.FirstSection.Body.Paragraphs)
            {
                Node importedNode = importer.ImportNode(paragraph, true);
                dstDoc.FirstSection.Body.AppendChild(importedNode);
            }

            dstDoc.Save(ArtifactsDir + "DocumentBuilder.IgnoreTextBoxes.docx");
            //ExEnd

            dstDoc = new Document(ArtifactsDir + "DocumentBuilder.IgnoreTextBoxes.docx");
            textBox = (Shape)dstDoc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual("Textbox contents", textBox.GetText().Trim());

            if (isIgnoreTextBoxes)
            {
                Assert.AreEqual(12.0d, textBox.FirstParagraph.Runs[0].Font.Size);
                Assert.AreEqual("Times New Roman", textBox.FirstParagraph.Runs[0].Font.Name);
            }
            else
            {
                Assert.AreEqual(24.0d, textBox.FirstParagraph.Runs[0].Font.Size);
                Assert.AreEqual("Courier New", textBox.FirstParagraph.Runs[0].Font.Name);
            }
        }

        [Test]
        public void MoveToField()
        {
            //ExStart
            //ExFor:DocumentBuilder.MoveToField
            //ExSummary:Shows how to move document builder's cursor to a specific field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a field using the DocumentBuilder and add a run of text after it
            Field field = builder.InsertField("MERGEFIELD field");
            builder.Write(" Text after the field.");

            // The builder's cursor is currently at end of the document
            Assert.Null(builder.CurrentNode);

            // We can move the builder to a field like this, placing the cursor at immediately after the field
            builder.MoveToField(field, true);

            // Note that the cursor is at a place past the FieldEnd node of the field, meaning that we are not actually inside the field
            // If we wish to move the DocumentBuilder to inside a field,
            // we will need to move it to a field's FieldStart or FieldSeparator node using the DocumentBuilder.MoveTo() method
            Assert.AreEqual(field.End, builder.CurrentNode.PreviousSibling);

            builder.Write(" Text immediately after the field.");
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);

            Assert.AreEqual("\u0013MERGEFIELD field\u0014«field»\u0015 Text immediately after the field. Text after the field.", doc.GetText().Trim());
        }

        [Test]
        public void InsertOleObjectException()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Assert.That(() => builder.InsertOleObject("", "checkbox", false, true, null),
                Throws.TypeOf<ArgumentException>());
        }

        [Test]
        public void InsertChartDouble()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertChart(ChartType, Double, Double)
            //ExSummary:Shows how to insert a chart into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertChart(ChartType.Pie, ConvertUtil.PixelToPoint(300), ConvertUtil.PixelToPoint(300));
            Assert.AreEqual(225.0d, ConvertUtil.PixelToPoint(300)); //ExSkip

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertedChartDouble.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertedChartDouble.docx");
            Shape chartShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual("Chart Title", chartShape.Chart.Title.Text);
            Assert.AreEqual(225.0d, chartShape.Width);
            Assert.AreEqual(225.0d, chartShape.Height);
        }

        [Test]
        public void InsertChartRelativePosition()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertChart(ChartType, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows how to insert a chart into a document and specify position and size.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertChart(ChartType.Pie, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin,
                100, 200, 100, WrapType.Square);

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertedChartRelativePosition.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertedChartRelativePosition.docx");
            Shape chartShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.AreEqual(100.0d, chartShape.Top);
            Assert.AreEqual(100.0d, chartShape.Left);
            Assert.AreEqual(200.0d, chartShape.Width);
            Assert.AreEqual(100.0d, chartShape.Height);
            Assert.AreEqual(WrapType.Square, chartShape.WrapType);
            Assert.AreEqual(RelativeHorizontalPosition.Margin, chartShape.RelativeHorizontalPosition);
            Assert.AreEqual(RelativeVerticalPosition.Margin, chartShape.RelativeVerticalPosition);
        }

        [Test]
        public void InsertField()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertField(String)
            //ExFor:Field
            //ExFor:Field.Update
            //ExFor:Field.Result
            //ExFor:Field.GetFieldCode
            //ExFor:Field.Type
            //ExFor:Field.Remove
            //ExFor:FieldType
            //ExSummary:Shows how to insert a field into a document by FieldCode.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a simple Date field into the document
            // When we insert a field through the DocumentBuilder class we can get the
            // special Field object which contains information about the field
            Field dateField = builder.InsertField(@"DATE \* MERGEFORMAT");

            // Update this particular field in the document so we can get the FieldResult
            dateField.Update();

            // Display some information from this field
            // The field result is where the last evaluated value is stored. This is what is displayed in the document
            // When field codes are not showing
            Assert.AreEqual(DateTime.Today, DateTime.Parse(dateField.Result));

            // Display the field code which defines the behavior of the field. This can been seen in Microsoft Word by pressing ALT+F9
            Assert.AreEqual(@"DATE \* MERGEFORMAT", dateField.GetFieldCode());

            // The field type defines what type of field in the Document this is. In this case the type is "FieldDate" 
            Assert.AreEqual(FieldType.FieldDate, dateField.Type);

            // Finally let's completely remove the field from the document. This can easily be done by invoking the Remove method on the object
            dateField.Remove();
            //ExEnd			

            Assert.AreEqual(0, doc.Range.Fields.Count);
        }

        [Test]
        public void InsertFieldByType()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertField(FieldType, Boolean)
            //ExSummary:Shows how to insert a field into a document using FieldType.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert an AUTHOR field using a DocumentBuilder
            doc.BuiltInDocumentProperties.Author = "John Doe";
            builder.Write("This document was written by ");
            builder.InsertField(FieldType.FieldAuthor, true);
            Assert.AreEqual(" AUTHOR ", doc.Range.Fields[0].GetFieldCode()); //ExSkip
            Assert.AreEqual("John Doe", doc.Range.Fields[0].Result); //ExSkip

            // Insert a PAGE field using a DocumentBuilder, but do not immediately update it
            builder.Write("\nThis is page ");
            builder.InsertField(FieldType.FieldPage, false);
            Assert.AreEqual(" PAGE ", doc.Range.Fields[1].GetFieldCode()); //ExSkip
            Assert.AreEqual("", doc.Range.Fields[1].Result); //ExSkip
            
            // Some fields types, such as ones that display document word/page counts may not keep track of their results in real time,
            // and will only display an accurate result during a field update
            // We can defer the updating of those fields until right before we need to see an accurate result
            // This method will manually update all the fields in a document
            doc.UpdateFields();

            Assert.AreEqual("1", doc.Range.Fields[1].Result);
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);

            Assert.AreEqual("This document was written by \u0013 AUTHOR \u0014John Doe\u0015" +
                            "\rThis is page \u0013 PAGE \u00141\u0015", doc.GetText().Trim());

            TestUtil.VerifyField(FieldType.FieldAuthor, " AUTHOR ", "John Doe", doc.Range.Fields[0]);
            TestUtil.VerifyField(FieldType.FieldPage, " PAGE ", "1", doc.Range.Fields[1]);
        }

        //ExStart
        //ExFor:IFieldResultFormatter
        //ExFor:IFieldResultFormatter.Format(Double, GeneralFormat)
        //ExFor:IFieldResultFormatter.Format(String, GeneralFormat)
        //ExFor:IFieldResultFormatter.FormatDateTime(DateTime, String, CalendarType)
        //ExFor:IFieldResultFormatter.FormatNumeric(Double, String)
        //ExFor:FieldOptions.ResultFormatter
        //ExFor:CalendarType
        //ExSummary:Shows how to control how the field result is formatted.
        [Test] //ExSkip
        public void FieldResultFormatting()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            doc.FieldOptions.ResultFormatter = new FieldResultFormatter("${0}", "Date: {0}", "Item # {0}:");

            // Insert a field with a numeric format
            builder.InsertField(" = 2 + 3 \\# $###", null);

            // Insert a field with a date/time format
            builder.InsertField("DATE \\@ \"d MMMM yyyy\"", null);

            // Insert a field with a general format
            builder.InsertField("QUOTE \"2\" \\* Ordinal", null);

            // Formats will be applied and recorded by the formatter during the field update
            doc.UpdateFields();
            ((FieldResultFormatter)doc.FieldOptions.ResultFormatter).PrintInvocations();

            // Our formatter has also overridden the formats that were originally applied in the fields
            Assert.AreEqual("$5", doc.Range.Fields[0].Result);
            Assert.IsTrue(doc.Range.Fields[1].Result.StartsWith("Date: "));
            Assert.AreEqual("Item # 2:", doc.Range.Fields[2].Result);
        }

        /// <summary>
        /// Custom IFieldResult implementation that applies formats and tracks format invocations
        /// </summary>
        private class FieldResultFormatter : IFieldResultFormatter
        {
            public FieldResultFormatter(string numberFormat, string dateFormat, string generalFormat)
            {
                mNumberFormat = numberFormat;
                mDateFormat = dateFormat;
                mGeneralFormat = generalFormat;
            }

            public string FormatNumeric(double value, string format)
            {
                mNumberFormatInvocations.Add(new object[] { value, format });

                return string.IsNullOrEmpty(mNumberFormat) ? null : string.Format(mNumberFormat, value);
            }

            public string FormatDateTime(DateTime value, string format, CalendarType calendarType)
            {
                mDateFormatInvocations.Add(new object[] { value, format, calendarType });

                return string.IsNullOrEmpty(mDateFormat) ? null : string.Format(mDateFormat, value);
            }

            public string Format(string value, GeneralFormat format)
            {
                return Format((object)value, format);
            }

            public string Format(double value, GeneralFormat format)
            {
                return Format((object)value, format);
            }

            private string Format(object value, GeneralFormat format)
            {
                mGeneralFormatInvocations.Add(new object[] { value, format });

                return string.IsNullOrEmpty(mGeneralFormat) ? null : string.Format(mGeneralFormat, value);
            }

            public void PrintInvocations()
            {
                Console.WriteLine("Number format invocations ({0}):", mNumberFormatInvocations.Count);
                foreach (object[] s in mNumberFormatInvocations)
                {
                    Console.WriteLine("\tValue: " + s[0] + ", original format: " + s[1]);
                }

                Console.WriteLine("Date format invocations ({0}):", mDateFormatInvocations.Count);
                foreach (object[] s in mDateFormatInvocations)
                {
                    Console.WriteLine("\tValue: " + s[0] + ", original format: " + s[1] + ", calendar type: " + s[2]);
                }

                Console.WriteLine("General format invocations ({0}):", mGeneralFormatInvocations.Count);
                foreach (object[] s in mGeneralFormatInvocations)
                {
                    Console.WriteLine("\tValue: " + s[0] + ", original format: " + s[1]);
                }
            }

            private readonly string mNumberFormat;
            private readonly string mDateFormat;
            private readonly string mGeneralFormat;

            private readonly ArrayList mNumberFormatInvocations = new ArrayList();
            private readonly ArrayList mDateFormatInvocations = new ArrayList();
            private readonly ArrayList mGeneralFormatInvocations = new ArrayList();

        }
        //ExEnd

        [Test]
        public void InsertVideoWithUrl()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertOnlineVideo(String, Double, Double)
            //ExSummary:Shows how to insert online video into a document using video url
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a video from Youtube
            builder.InsertOnlineVideo("https://youtu.be/t_1LYZ102RA", 360, 270);

            // Click on the shape in the output document to watch the video from Microsoft Word
            doc.Save(ArtifactsDir + "DocumentBuilder.InsertVideoWithUrl.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertVideoWithUrl.docx");
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(480, 360, ImageType.Jpeg, shape);
            TestUtil.VerifyWebResponseStatusCode(HttpStatusCode.OK, shape.HRef);

            Assert.AreEqual(360.0d, shape.Width);
            Assert.AreEqual(270.0d, shape.Height);
        }

        [Test]
        public void InsertUnderline()
        {
            //ExStart
            //ExFor:DocumentBuilder.Underline
            //ExSummary:Shows how to set and edit a document builder's underline.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set a new style for our underline
            builder.Underline = Underline.Dash;

            // Same object as DocumentBuilder.Font.Underline
            Assert.AreEqual(builder.Underline, builder.Font.Underline);
            Assert.AreEqual(Underline.Dash, builder.Font.Underline);

            // These properties will be applied to the underline as well
            builder.Font.Color = Color.Blue;
            builder.Font.Size = 32;

            builder.Writeln("Underlined text.");

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertUnderline.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertUnderline.docx");
            Run firstRun = doc.FirstSection.Body.FirstParagraph.Runs[0];

            Assert.AreEqual("Underlined text.", firstRun.GetText().Trim());
            Assert.AreEqual(Underline.Dash, firstRun.Font.Underline);
            Assert.AreEqual(Color.Blue.ToArgb(), firstRun.Font.Color.ToArgb());
            Assert.AreEqual(32.0d, firstRun.Font.Size);
        }

        [Test]
        public void CurrentStory()
        {
            //ExStart
            //ExFor:DocumentBuilder.CurrentStory
            //ExSummary:Shows how to work with a document builder's current story.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // A Story is a type of node that have child Paragraph nodes, such as a Body,
            // which would usually be a parent node to a DocumentBuilder's current paragraph
            Assert.AreEqual(builder.CurrentStory, doc.FirstSection.Body);
            Assert.AreEqual(builder.CurrentStory, builder.CurrentParagraph.ParentNode);
            Assert.AreEqual(StoryType.MainText, builder.CurrentStory.StoryType);

            builder.CurrentStory.AppendParagraph("Text added to current Story.");

            // A Story can contain tables too
            Table table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Row 1 cell 1");
            builder.InsertCell();
            builder.Write("Row 1 cell 2");
            builder.EndTable();

            // The table we just made is automatically placed in the story
            Assert.IsTrue(builder.CurrentStory.Tables.Contains(table));
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            Assert.AreEqual(1, doc.FirstSection.Body.Tables.Count);
            Assert.AreEqual("Row 1 cell 1\aRow 1 cell 2\a\a\rText added to current Story.", doc.FirstSection.Body.GetText().Trim());
        }

        [Test]
        public void InsertOlePowerpoint()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertOleObject(Stream, String, Boolean, Image)
            //ExSummary:Shows how to use document builder to embed Ole objects in a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Let's take a spreadsheet from our system and insert it into the document
            using (Stream spreadsheetStream = File.Open(MyDir + "Spreadsheet.xlsx", FileMode.Open))
            {
                // The spreadsheet can be activated by double clicking the panel that you'll see in the document immediately under the text we will add
                // We did not set the area to double click as an icon nor did we change its appearance so it looks like a simple panel
                builder.Writeln("Spreadsheet Ole object:");
                builder.InsertOleObject(spreadsheetStream, "OleObject.xlsx", false, null);

                // A powerpoint presentation is another type of object we can embed in our document
                // This time we'll also exercise some control over how it looks 
                using (Stream powerpointStream = File.Open(MyDir + "Presentation.pptx", FileMode.Open))
                {
                    // If we insert the Ole object as an icon, we are still provided with a default icon
                    // If that is not suitable, we can make the icon to look like any image
                    using (WebClient webClient = new WebClient())
                    {
                        byte[] imgBytes = webClient.DownloadData(AsposeLogoUrl);

                        #if NETCOREAPP2_1 || __MOBILE__
                        
                        SKBitmap bitmap = SKBitmap.Decode(imgBytes);
                        builder.InsertParagraph();
                        builder.Writeln("Powerpoint Ole object:");
                        builder.InsertOleObject(powerpointStream, "MyOleObject.pptx", true, bitmap);
                        
                        #elif NET462
                        
                        using (MemoryStream stream = new MemoryStream(imgBytes))
                        {
                            using (Image image = Image.FromStream(stream))
                            {
                                // If we double click the image, the powerpoint presentation will open
                                builder.InsertParagraph();
                                builder.Writeln("Powerpoint Ole object:");
                                builder.InsertOleObject(powerpointStream, "OleObject.pptx", true, image);
                            }
                        }

                        #endif
                    }
                }
            }

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertOlePowerpoint.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertOlePowerpoint.docx");

            Assert.AreEqual(2, doc.GetChildNodes(NodeType.Shape, true).Count);

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Assert.AreEqual("", shape.OleFormat.IconCaption);
            Assert.False(shape.OleFormat.OleIcon);

            shape = (Shape)doc.GetChild(NodeType.Shape, 1, true);
            Assert.AreEqual("Unknown", shape.OleFormat.IconCaption);
            Assert.True(shape.OleFormat.OleIcon);
        }

        [Test]
        public void InsertStyleSeparator()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertStyleSeparator
            //ExSummary:Shows how to separate styles from two different paragraphs used in one logical printed paragraph.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Append text in the "Heading 1" style
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Write("This text is in a Heading style. ");

            // Insert a style separator
            builder.InsertStyleSeparator();

            // The style separator appears in the form of a paragraph break that doesn't start a new line
            // So, while this looks like one continuous paragraph with two styles in the output document, 
            // it is actually two paragraphs with different styles, but no line break between the first and second paragraph
            Assert.AreEqual(2, doc.FirstSection.Body.Paragraphs.Count);

            // Append text with another style
            Style paraStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MyParaStyle");
            paraStyle.Font.Bold = false;
            paraStyle.Font.Size = 8;
            paraStyle.Font.Name = "Arial";

            // Set the style of the current paragraph to our custom style
            // This will apply to only the text after the style separator
            builder.ParagraphFormat.StyleName = paraStyle.Name;
            builder.Write("This text is in a custom style. ");

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertStyleSeparator.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertStyleSeparator.docx");

            Assert.AreEqual(2, doc.FirstSection.Body.Paragraphs.Count);
            Assert.AreEqual("This text is in a Heading style. \r This text is in a custom style.",
                doc.GetText().Trim());
        }

        [Test]
        public void WithoutStyleSeparator()
        {
            DocumentBuilder builder = new DocumentBuilder(new Document());

            Style paraStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MyParaStyle");
            paraStyle.Font.Bold = false;
            paraStyle.Font.Size = 8;
            paraStyle.Font.Name = "Arial";

            // Append text with "Heading 1" style
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Write("This text is in a Heading style. ");

            // Append text with another style
            builder.ParagraphFormat.StyleName = paraStyle.Name;
            builder.Write("This text is in a custom style. ");

            builder.Document.Save(ArtifactsDir + "DocumentBuilder.WithoutStyleSeparator.docx");
        }

        [Test]
        public void SmartStyleBehavior()
        {
            //ExStart
            //ExFor:ImportFormatOptions
            //ExFor:ImportFormatOptions.SmartStyleBehavior
            //ExFor:DocumentBuilder.InsertDocument(Document, ImportFormatMode, ImportFormatOptions)
            //ExSummary:Shows how to resolve styles behavior while inserting documents.
            Document dstDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(dstDoc);

            Style myStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MyStyle");
            myStyle.Font.Size = 14;
            myStyle.Font.Name = "Courier New";
            myStyle.Font.Color = Color.Blue;

            // Append text with custom style
            builder.ParagraphFormat.StyleName = myStyle.Name;
            builder.Writeln("Hello world!");

            // Clone the document, and edit the clone's "MyStyle" style so it is a different color than that of the original
            // If we append this document to the original, the different styles will clash since they are the same name, and we will need to resolve it
            Document srcDoc = dstDoc.Clone();
            srcDoc.Styles["MyStyle"].Font.Color = Color.Red;

            // When SmartStyleBehavior is enabled,
            // a source style will be expanded into a direct attributes inside a destination document,
            // if KeepSourceFormatting importing mode is used
            ImportFormatOptions options = new ImportFormatOptions();
            options.SmartStyleBehavior = true;

            builder.InsertDocument(srcDoc, ImportFormatMode.KeepSourceFormatting, options);

            dstDoc.Save(ArtifactsDir + "DocumentBuilder.SmartStyleBehavior.docx");
            //ExEnd

            dstDoc = new Document(ArtifactsDir + "DocumentBuilder.SmartStyleBehavior.docx");

            Assert.AreEqual(Color.Blue.ToArgb(), dstDoc.Styles["MyStyle"].Font.Color.ToArgb());
            Assert.AreEqual("MyStyle", dstDoc.FirstSection.Body.Paragraphs[0].ParagraphFormat.Style.Name);

            Assert.AreEqual("Normal", dstDoc.FirstSection.Body.Paragraphs[1].ParagraphFormat.Style.Name);
            Assert.AreEqual(14, dstDoc.FirstSection.Body.Paragraphs[1].Runs[0].Font.Size);
            Assert.AreEqual("Courier New", dstDoc.FirstSection.Body.Paragraphs[1].Runs[0].Font.Name);
            Assert.AreEqual(Color.Red.ToArgb(), dstDoc.FirstSection.Body.Paragraphs[1].Runs[0].Font.Color.ToArgb());
        }

        [Test]
        public void IgnoreHeaderFooter()
        {
            //ExStart
            //ExFor:ImportFormatOptions.IgnoreHeaderFooter
            //ExSummary:Shows how to specifies ignoring source formatting of headers/footers content.
            Document dstDoc = new Document(MyDir + "Document.docx");
            Document srcDoc = new Document(MyDir + "Header and footer types.docx");
 
            ImportFormatOptions importFormatOptions = new ImportFormatOptions();
            importFormatOptions.IgnoreHeaderFooter = true;
 
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting, importFormatOptions);

            dstDoc.Save(ArtifactsDir + "DocumentBuilder.IgnoreHeaderFooter.docx");
            //ExEnd
        }

        #if NET462 || NETCOREAPP2_1 || JAVA
        /// <summary>
        /// All markdown tests work with the same file
        /// That's why we need order for them 
        /// </summary>
        [Test, Order(1), Category("SkipTearDown")]
        public void MarkdownDocumentEmphases()
        {
            DocumentBuilder builder = new DocumentBuilder();
            
            // Bold and Italic are represented as Font.Bold and Font.Italic
            builder.Font.Italic = true;
            builder.Writeln("This text will be italic");
            
            // Use clear formatting if don't want to combine styles between paragraphs
            builder.Font.ClearFormatting();
            
            builder.Font.Bold = true;
            builder.Writeln("This text will be bold");
            
            builder.Font.ClearFormatting();
            
            // You can also write create BoldItalic text
            builder.Font.Italic = true;
            builder.Write("You ");
            builder.Font.Bold = true;
            builder.Write("can");
            builder.Font.Bold = false;
            builder.Writeln(" combine them");

            builder.Font.ClearFormatting();

            builder.Font.StrikeThrough = true;
            builder.Writeln("This text will be strikethrough");
            
            // Markdown treats asterisks (*), underscores (_) and tilde (~) as indicators of emphasis
            builder.Document.Save(ArtifactsDir + "DocumentBuilder.MarkdownDocument.md");
        }

        /// <summary>
        /// All markdown tests work with the same file
        /// That's why we need order for them 
        /// </summary>
        [Test, Order(2), Category("SkipTearDown")]
        public void MarkdownDocumentInlineCode()
        {
            Document doc = new Document(ArtifactsDir + "DocumentBuilder.MarkdownDocument.md");
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            // Prepare our created document for further work
            // And clear paragraph formatting not to use the previous styles
            builder.MoveToDocumentEnd();
            builder.ParagraphFormat.ClearFormatting();
            builder.Writeln("\n");
            
            // Style with name that starts from word InlineCode, followed by optional dot (.) and number of backticks (`)
            // If number of backticks is missed, then one backtick will be used by default
            Style inlineCode1BackTicks = doc.Styles.Add(StyleType.Character, "InlineCode");
            builder.Font.Style = inlineCode1BackTicks;
            builder.Writeln("Text with InlineCode style with one backtick");
            
            // Use optional dot (.) and number of backticks (`)
            // There will be 3 backticks
            Style inlineCode3BackTicks = doc.Styles.Add(StyleType.Character, "InlineCode.3");
            builder.Font.Style = inlineCode3BackTicks;
            builder.Writeln("Text with InlineCode style with 3 backticks");

            builder.Document.Save(ArtifactsDir + "DocumentBuilder.MarkdownDocument.md");
        }

        /// <summary>
        /// All markdown tests work with the same file
        /// That's why we need order for them 
        /// </summary>
        [Test, Order(3), Category("SkipTearDown")]
        [Description("WORDSNET-19850")]
        public void MarkdownDocumentHeadings()
        {
            Document doc = new Document(ArtifactsDir + "DocumentBuilder.MarkdownDocument.md");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Prepare our created document for further work
            // And clear paragraph formatting not to use the previous styles
            builder.MoveToDocumentEnd();
            builder.ParagraphFormat.ClearFormatting();
            builder.Writeln("\n");
            
            // By default Heading styles in Word may have bold and italic formatting
            // If we do not want text to be emphasized, set these properties explicitly to false
            // Thus we can't use 'builder.Font.ClearFormatting()' because Bold/Italic will be set to true
            builder.Font.Bold = false;
            builder.Font.Italic = false;
            
            // Create for one heading for each level
            builder.ParagraphFormat.StyleName = "Heading 1";
            builder.Font.Italic = true;
            builder.Writeln("This is an italic H1 tag");

            // Reset our styles from the previous paragraph to not combine styles between paragraphs
            builder.Font.Bold = false;
            builder.Font.Italic = false;

            // Structure-enhanced text heading can be added through style inheritance
            Style setextHeading1 = doc.Styles.Add(StyleType.Paragraph, "SetextHeading1");
            builder.ParagraphFormat.Style = setextHeading1;
            doc.Styles["SetextHeading1"].BaseStyleName = "Heading 1";
            builder.Writeln("SetextHeading 1");
            
            builder.ParagraphFormat.StyleName = "Heading 2";
            builder.Writeln("This is an H2 tag");

            builder.Font.Bold = false;
            builder.Font.Italic = false;

            Style setextHeading2 = doc.Styles.Add(StyleType.Paragraph, "SetextHeading2");
            builder.ParagraphFormat.Style = setextHeading2;
            doc.Styles["SetextHeading2"].BaseStyleName = "Heading 2";
            builder.Writeln("SetextHeading 2");
            
            builder.ParagraphFormat.Style = doc.Styles["Heading 3"];
            builder.Writeln("This is an H3 tag");
            
            builder.Font.Bold = false;
            builder.Font.Italic = false;

            builder.ParagraphFormat.Style = doc.Styles["Heading 4"];
            builder.Font.Bold = true;
            builder.Writeln("This is an bold H4 tag");
            
            builder.Font.Bold = false;
            builder.Font.Italic = false;

            builder.ParagraphFormat.Style = doc.Styles["Heading 5"];
            builder.Font.Italic = true;
            builder.Font.Bold = true;
            builder.Writeln("This is an italic and bold H5 tag");
            
            builder.Font.Bold = false;
            builder.Font.Italic = false;

            builder.ParagraphFormat.Style = doc.Styles["Heading 6"];
            builder.Writeln("This is an H6 tag");
            
            doc.Save(ArtifactsDir + "DocumentBuilder.MarkdownDocument.md");
        }

        /// <summary>
        /// All markdown tests work with the same file
        /// That's why we need order for them 
        /// </summary>
        [Test, Order(4), Category("SkipTearDown")]
        public void MarkdownDocumentBlockquotes()
        {
            Document doc = new Document(ArtifactsDir + "DocumentBuilder.MarkdownDocument.md");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Prepare our created document for further work
            // And clear paragraph formatting not to use the previous styles
            builder.MoveToDocumentEnd();
            builder.ParagraphFormat.ClearFormatting();
            builder.Writeln("\n");

            // By default document stores blockquote style for the first level
            builder.ParagraphFormat.StyleName = "Quote";
            builder.Writeln("Blockquote");
            
            // Create styles for nested levels through style inheritance
            Style quoteLevel2 = doc.Styles.Add(StyleType.Paragraph, "Quote1");
            builder.ParagraphFormat.Style = quoteLevel2;
            doc.Styles["Quote1"].BaseStyleName = "Quote";
            builder.Writeln("1. Nested blockquote");
            
            Style quoteLevel3 = doc.Styles.Add(StyleType.Paragraph, "Quote2");
            builder.ParagraphFormat.Style = quoteLevel3;
            doc.Styles["Quote2"].BaseStyleName = "Quote1";
            builder.Font.Italic = true;
            builder.Writeln("2. Nested italic blockquote");
            
            Style quoteLevel4 = doc.Styles.Add(StyleType.Paragraph, "Quote3");
            builder.ParagraphFormat.Style = quoteLevel4;
            doc.Styles["Quote3"].BaseStyleName = "Quote2";
            builder.Font.Italic = false;
            builder.Font.Bold = true;
            builder.Writeln("3. Nested bold blockquote");
            
            Style quoteLevel5 = doc.Styles.Add(StyleType.Paragraph, "Quote4");
            builder.ParagraphFormat.Style = quoteLevel5;
            doc.Styles["Quote4"].BaseStyleName = "Quote3";
            builder.Font.Bold = false;
            builder.Writeln("4. Nested blockquote");
            
            Style quoteLevel6 = doc.Styles.Add(StyleType.Paragraph, "Quote5");
            builder.ParagraphFormat.Style = quoteLevel6;
            doc.Styles["Quote5"].BaseStyleName = "Quote4";
            builder.Writeln("5. Nested blockquote");
            
            Style quoteLevel7 = doc.Styles.Add(StyleType.Paragraph, "Quote6");
            builder.ParagraphFormat.Style = quoteLevel7;
            doc.Styles["Quote6"].BaseStyleName = "Quote5";
            builder.Font.Italic = true;
            builder.Font.Bold = true;
            builder.Writeln("6. Nested italic bold blockquote");
            
            doc.Save(ArtifactsDir + "DocumentBuilder.MarkdownDocument.md");
        }

        /// <summary>
        /// All markdown tests work with the same file
        /// That's why we need order for them 
        /// </summary>
        [Test, Order(5), Category("SkipTearDown")]
        public void MarkdownDocumentIndentedCode()
        {
            Document doc = new Document(ArtifactsDir + "DocumentBuilder.MarkdownDocument.md");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Prepare our created document for further work
            // And clear paragraph formatting not to use the previous styles
            builder.MoveToDocumentEnd();
            builder.Writeln("\n");
            builder.ParagraphFormat.ClearFormatting();
            builder.Writeln("\n");

            Style indentedCode = doc.Styles.Add(StyleType.Paragraph, "IndentedCode");
            builder.ParagraphFormat.Style = indentedCode;
            builder.Writeln("This is an indented code");
            
            doc.Save(ArtifactsDir + "DocumentBuilder.MarkdownDocument.md");
        }

        /// <summary>
        /// All markdown tests work with the same file
        /// That's why we need order for them 
        /// </summary>
        [Test, Order(6), Category("SkipTearDown")]
        public void MarkdownDocumentFencedCode()
        {
            Document doc = new Document(ArtifactsDir + "DocumentBuilder.MarkdownDocument.md");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Prepare our created document for further work
            // And clear paragraph formatting not to use the previous styles
            builder.MoveToDocumentEnd();
            builder.Writeln("\n");
            builder.ParagraphFormat.ClearFormatting();
            builder.Writeln("\n");

            Style fencedCode = doc.Styles.Add(StyleType.Paragraph, "FencedCode");
            builder.ParagraphFormat.Style = fencedCode;
            builder.Writeln("This is a fenced code");

            Style fencedCodeWithInfo = doc.Styles.Add(StyleType.Paragraph, "FencedCode.C#");
            builder.ParagraphFormat.Style = fencedCodeWithInfo;
            builder.Writeln("This is a fenced code with info string");

            doc.Save(ArtifactsDir + "DocumentBuilder.MarkdownDocument.md");
        }

        /// <summary>
        /// All markdown tests work with the same file
        /// That's why we need order for them 
        /// </summary>
        [Test, Order(7), Category("SkipTearDown")]
        public void MarkdownDocumentHorizontalRule()
        {
            Document doc = new Document(ArtifactsDir + "DocumentBuilder.MarkdownDocument.md");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Prepare our created document for further work
            // And clear paragraph formatting not to use the previous styles
            builder.MoveToDocumentEnd();
            builder.ParagraphFormat.ClearFormatting();
            builder.Writeln("\n");

            // Insert HorizontalRule that will be present in .md file as '-----'
            builder.InsertHorizontalRule();
 
            builder.Document.Save(ArtifactsDir + "DocumentBuilder.MarkdownDocument.md");
        }

        /// <summary>
        /// All markdown tests work with the same file
        /// That's why we need order for them 
        /// </summary>
        [Test, Order(8), Category("SkipTearDown")]
        public void MarkdownDocumentBulletedList()
        {
            Document doc = new Document(ArtifactsDir + "DocumentBuilder.MarkdownDocument.md");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Prepare our created document for further work
            // And clear paragraph formatting not to use the previous styles
            builder.MoveToDocumentEnd();
            builder.ParagraphFormat.ClearFormatting();
            builder.Writeln("\n");

            // Bulleted lists are represented using paragraph numbering
            builder.ListFormat.ApplyBulletDefault();
            // There can be 3 types of bulleted lists
            // The only diff in a numbering format of the very first level are: ‘-’, ‘+’ or ‘*’ respectively
            builder.ListFormat.List.ListLevels[0].NumberFormat = "-";
            
            builder.Writeln("Item 1");
            builder.Writeln("Item 2");
            builder.ListFormat.ListIndent();
            builder.Writeln("Item 2a");
            builder.Writeln("Item 2b");
 
            builder.Document.Save(ArtifactsDir + "DocumentBuilder.MarkdownDocument.md");
        }

        /// <summary>
        /// All markdown tests work with the same file.
        /// That's why we need order for them.
        /// </summary>
        [Test, Order(9)]
        [TestCase("Italic", "Normal", true, false, Category = "SkipTearDown")]
        [TestCase("Bold", "Normal", false, true, Category = "SkipTearDown")]
        [TestCase("ItalicBold", "Normal", true, true, Category = "SkipTearDown")]
        [TestCase("Text with InlineCode style with one backtick", "InlineCode", false, false, Category = "SkipTearDown")]
        [TestCase("Text with InlineCode style with 3 backticks", "InlineCode.3", false, false, Category = "SkipTearDown")]
        [TestCase("This is an italic H1 tag", "Heading 1", true, false, Category = "SkipTearDown")]
        [TestCase("SetextHeading 1", "SetextHeading1", false, false, Category = "SkipTearDown")]
        [TestCase("This is an H2 tag", "Heading 2", false, false, Category = "SkipTearDown")]
        [TestCase("SetextHeading 2", "SetextHeading2", false, false, Category = "SkipTearDown")]
        [TestCase("This is an H3 tag", "Heading 3", false, false, Category = "SkipTearDown")]
        [TestCase("This is an bold H4 tag", "Heading 4", false, true, Category = "SkipTearDown")]
        [TestCase("This is an italic and bold H5 tag", "Heading 5", true, true, Category = "SkipTearDown")]
        [TestCase("This is an H6 tag", "Heading 6", false, false, Category = "SkipTearDown")]
        [TestCase("Blockquote", "Quote", false, false, Category = "SkipTearDown")]
        [TestCase("1. Nested blockquote", "Quote1", false, false, Category = "SkipTearDown")]
        [TestCase("2. Nested italic blockquote", "Quote2", true, false, Category = "SkipTearDown")]
        [TestCase("3. Nested bold blockquote", "Quote3", false, true, Category = "SkipTearDown")]
        [TestCase("4. Nested blockquote", "Quote4", false, false, Category = "SkipTearDown")]
        [TestCase("5. Nested blockquote", "Quote5", false, false, Category = "SkipTearDown")]
        [TestCase("6. Nested italic bold blockquote", "Quote6", true, true, Category = "SkipTearDown")]
        [TestCase("This is an indented code", "IndentedCode", false, false, Category = "SkipTearDown")]
        [TestCase("This is a fenced code", "FencedCode", false, false, Category = "SkipTearDown")]
        [TestCase("This is a fenced code with info string", "FencedCode.C#", false, false, Category = "SkipTearDown")]
        [TestCase("Item 1", "Normal", false, false)]
        public void LoadMarkdownDocumentAndAssertContent(string text, string styleName, bool isItalic, bool isBold)
        {
            // Load created document from previous tests
            Document doc = new Document(ArtifactsDir + "DocumentBuilder.MarkdownDocument.md");
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;
            
            foreach (Paragraph paragraph in paragraphs)
            {
                if (paragraph.Runs.Count != 0)
                {
                    // Check that all document text has the necessary styles
                    if (paragraph.Runs[0].Text == text && !text.Contains("InlineCode"))
                    {
                        Assert.AreEqual(styleName, paragraph.ParagraphFormat.Style.Name);
                        Assert.AreEqual(isItalic, paragraph.Runs[0].Font.Italic);
                        Assert.AreEqual(isBold, paragraph.Runs[0].Font.Bold);
                    }
                    else if (paragraph.Runs[0].Text == text && text.Contains("InlineCode"))
                    {
                        Assert.AreEqual(styleName, paragraph.Runs[0].Font.StyleName);
                    }
                }

                // Check that document also has a HorizontalRule present as a shape
                NodeCollection shapesCollection = doc.FirstSection.Body.GetChildNodes(NodeType.Shape, true);
                Shape horizontalRuleShape = (Shape) shapesCollection[0];
                
                Assert.IsTrue(shapesCollection.Count == 1);
                Assert.IsTrue(horizontalRuleShape.IsHorizontalRule);
            }
        }

        [TestCase(TableContentAlignment.Left)]
        [TestCase(TableContentAlignment.Right)]
        [TestCase(TableContentAlignment.Center)]
        [TestCase(TableContentAlignment.Auto)]
        public void MarkdownDocumentTableContentAlignment(TableContentAlignment tableContentAlignment)
        {
            DocumentBuilder builder = new DocumentBuilder();

            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Right;
            builder.Write("Cell1");
            builder.InsertCell();
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Write("Cell2");

            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
            saveOptions.TableContentAlignment = tableContentAlignment;

            builder.Document.Save(ArtifactsDir + "MarkdownDocumentTableContentAlignment.md", saveOptions);

            Document doc = new Document(ArtifactsDir + "MarkdownDocumentTableContentAlignment.md");
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);

            switch (tableContentAlignment)
            {
                case TableContentAlignment.Auto:
                    Assert.AreEqual(ParagraphAlignment.Right,
                        table.FirstRow.Cells[0].FirstParagraph.ParagraphFormat.Alignment);
                    Assert.AreEqual(ParagraphAlignment.Center,
                        table.FirstRow.Cells[1].FirstParagraph.ParagraphFormat.Alignment);
                    break;
                case TableContentAlignment.Left:
                    Assert.AreEqual(ParagraphAlignment.Left,
                        table.FirstRow.Cells[0].FirstParagraph.ParagraphFormat.Alignment);
                    Assert.AreEqual(ParagraphAlignment.Left,
                        table.FirstRow.Cells[1].FirstParagraph.ParagraphFormat.Alignment);
                    break;
                case TableContentAlignment.Center:
                    Assert.AreEqual(ParagraphAlignment.Center,
                        table.FirstRow.Cells[0].FirstParagraph.ParagraphFormat.Alignment);
                    Assert.AreEqual(ParagraphAlignment.Center,
                        table.FirstRow.Cells[1].FirstParagraph.ParagraphFormat.Alignment);
                    break;
                case TableContentAlignment.Right:
                    Assert.AreEqual(ParagraphAlignment.Right,
                        table.FirstRow.Cells[0].FirstParagraph.ParagraphFormat.Alignment);
                    Assert.AreEqual(ParagraphAlignment.Right,
                        table.FirstRow.Cells[1].FirstParagraph.ParagraphFormat.Alignment);
                    break;
            }
        }

        [Test]
        public void InsertOnlineVideo()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertOnlineVideo(String, String, Byte[], Double, Double)
            //ExFor:DocumentBuilder.InsertOnlineVideo(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExFor:DocumentBuilder.InsertOnlineVideo(String, String, Byte[], RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows how to insert online video into a document using html code.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Visible url
            string vimeoVideoUrl = @"https://vimeo.com/52477838";

            // Embed Html code
            string vimeoEmbedCode =
                "<iframe src=\"https://player.vimeo.com/video/52477838\" width=\"640\" height=\"360\" frameborder=\"0\" title=\"Aspose\" webkitallowfullscreen mozallowfullscreen allowfullscreen></iframe>";

            // This video will have an automatically generated thumbnail, and we are setting the size according to its 16:9 aspect ratio
            builder.Writeln("Video with an automatically generated thumbnail at the top left corner of the page:");
            builder.InsertOnlineVideo(vimeoVideoUrl, RelativeHorizontalPosition.LeftMargin, 0,
                RelativeVerticalPosition.TopMargin, 0, 320, 180, WrapType.Square);
            builder.InsertBreak(BreakType.PageBreak);

            // We can get an image to use as a custom thumbnail
            using (WebClient webClient = new WebClient())
            {
                byte[] imageBytes = webClient.DownloadData(AsposeLogoUrl);

                using (MemoryStream stream = new MemoryStream(imageBytes))
                {
                    using (Image image = Image.FromStream(stream))
                    {
                        // This puts the video where we are with our document builder, with a custom thumbnail and size depending on the size of the image
                        builder.Writeln("Custom thumbnail at document builder's cursor:");
                        builder.InsertOnlineVideo(vimeoVideoUrl, vimeoEmbedCode, imageBytes, image.Width, image.Height);
                        builder.InsertBreak(BreakType.PageBreak);

                        // We can put the video at the bottom right edge of the page too, but we'll have to take the page margins into account 
                        double left = builder.PageSetup.RightMargin - image.Width;
                        double top = builder.PageSetup.BottomMargin - image.Height;

                        // Here we use a custom thumbnail and relative positioning to put it and the bottom right of tha page
                        builder.Writeln("Bottom right of page with custom thumbnail:");

                        builder.InsertOnlineVideo(vimeoVideoUrl, vimeoEmbedCode, imageBytes,
                            RelativeHorizontalPosition.RightMargin, left, RelativeVerticalPosition.BottomMargin, top,
                            image.Width, image.Height, WrapType.Square);
                    }
                }
            }

            doc.Save(ArtifactsDir + "DocumentBuilder.InsertOnlineVideo.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "DocumentBuilder.InsertOnlineVideo.docx");
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(640, 360, ImageType.Jpeg, shape);

            Assert.AreEqual(320.0d, shape.Width);
            Assert.AreEqual(180.0d, shape.Height);
            Assert.AreEqual(0.0d, shape.Left);
            Assert.AreEqual(0.0d, shape.Top);
            Assert.AreEqual(WrapType.Square, shape.WrapType);
            Assert.AreEqual(RelativeVerticalPosition.TopMargin, shape.RelativeVerticalPosition);
            Assert.AreEqual(RelativeHorizontalPosition.LeftMargin, shape.RelativeHorizontalPosition);

            Assert.AreEqual("https://vimeo.com/52477838", shape.HRef);

            shape = (Shape)doc.GetChild(NodeType.Shape, 1, true);

            TestUtil.VerifyImageInShape(320, 320, ImageType.Png, shape);
            Assert.AreEqual(320.0d, shape.Width);
            Assert.AreEqual(320.0d, shape.Height);
            Assert.AreEqual(0.0d, shape.Left);
            Assert.AreEqual(0.0d, shape.Top);
            Assert.AreEqual(WrapType.Inline, shape.WrapType);
            Assert.AreEqual(RelativeVerticalPosition.Paragraph, shape.RelativeVerticalPosition);
            Assert.AreEqual(RelativeHorizontalPosition.Column, shape.RelativeHorizontalPosition);

            Assert.AreEqual("https://vimeo.com/52477838", shape.HRef);

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            TestUtil.VerifyWebResponseStatusCode(HttpStatusCode.OK, shape.HRef);
        }
#endif
    }
}