using System;
using System.Drawing;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;
using NUnit.Framework;
using Font = Aspose.Words.Font;

namespace DocsExamples.Programming_with_Documents.Working_with_Document
{
    internal class AddContentUsingDocumentBuilder : DocsExamplesBase
    {
        [Test]
        public void DocumentBuilderInsertBookmark()
        {
            //ExStart:DocumentBuilderInsertBookmark
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartBookmark("FineBookmark");
            builder.Writeln("This is just a fine bookmark.");
            builder.EndBookmark("FineBookmark");

            doc.Save(ArtifactsDir + "WorkingWithBookmarks.DocumentBuilderInsertBookmark.docx");
            //ExEnd:DocumentBuilderInsertBookmark
        }

        [Test]
        public void BuildTable()
        {
            //ExStart:BuildTable
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.Write("This is row 1 cell 1");

            builder.InsertCell();
            builder.Write("This is row 1 cell 2");

            builder.EndRow();

            builder.InsertCell();
            
            builder.RowFormat.Height = 100;
            builder.RowFormat.HeightRule = HeightRule.Exactly;
            builder.CellFormat.Orientation = TextOrientation.Upward;
            builder.Writeln("This is row 2 cell 1");

            builder.InsertCell();
            builder.CellFormat.Orientation = TextOrientation.Downward;
            builder.Writeln("This is row 2 cell 2");

            builder.EndRow();
            builder.EndTable();

            doc.Save(ArtifactsDir + "AddContentUsingDocumentBuilder.BuildTable.docx");
            //ExEnd:BuildTable
        }

        [Test]
        public void InsertHorizontalRule()
        {
            //ExStart:InsertHorizontalRule
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Insert a horizontal rule shape into the document.");
            builder.InsertHorizontalRule();

            doc.Save(ArtifactsDir + "AddContentUsingDocumentBuilder.InsertHorizontalRule.docx");
            //ExEnd:InsertHorizontalRule
        }

        [Test]
        public void HorizontalRuleFormat()
        {
            //ExStart:HorizontalRuleFormat
            DocumentBuilder builder = new DocumentBuilder();

            Shape shape = builder.InsertHorizontalRule();
            
            HorizontalRuleFormat horizontalRuleFormat = shape.HorizontalRuleFormat;
            horizontalRuleFormat.Alignment = HorizontalRuleAlignment.Center;
            horizontalRuleFormat.WidthPercent = 70;
            horizontalRuleFormat.Height = 3;
            horizontalRuleFormat.Color = Color.Blue;
            horizontalRuleFormat.NoShade = true;

            builder.Document.Save(ArtifactsDir + "AddContentUsingDocumentBuilder.HorizontalRuleFormat.docx");
            //ExEnd:HorizontalRuleFormat
        }

        [Test]
        public void InsertBreak()
        {
            //ExStart:InsertBreak
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("This is page 1.");
            builder.InsertBreak(BreakType.PageBreak);

            builder.Writeln("This is page 2.");
            builder.InsertBreak(BreakType.PageBreak);

            builder.Writeln("This is page 3.");

            doc.Save(ArtifactsDir + "AddContentUsingDocumentBuilder.InsertBreak.docx");
            //ExEnd:InsertBreak
        }

        [Test]
        public void InsertTextInputFormField()
        {
            //ExStart:InsertTextInputFormField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Hello", 0);

            doc.Save(ArtifactsDir + "AddContentUsingDocumentBuilder.InsertTextInputFormField.docx");
            //ExEnd:InsertTextInputFormField
        }

        [Test]
        public void InsertCheckBoxFormField()
        {
            //ExStart:InsertCheckBoxFormField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.InsertCheckBox("CheckBox", true, true, 0);

            doc.Save(ArtifactsDir + "AddContentUsingDocumentBuilder.InsertCheckBoxFormField.docx");
            //ExEnd:InsertCheckBoxFormField
        }

        [Test]
        public void InsertComboBoxFormField()
        {
            //ExStart:InsertComboBoxFormField
            string[] items = { "One", "Two", "Three" };

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertComboBox("DropDown", items, 0);

            doc.Save(ArtifactsDir + "AddContentUsingDocumentBuilder.InsertComboBoxFormField.docx");
            //ExEnd:InsertComboBoxFormField
        }

        [Test]
        public void InsertHtml()
        {
            //ExStart:InsertHtml
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.InsertHtml(
                "<P align='right'>Paragraph right</P>" +
                "<b>Implicit paragraph left</b>" +
                "<div align='center'>Div center</div>" +
                "<h1 align='left'>Heading 1 left.</h1>");

            doc.Save(ArtifactsDir + "AddContentUsingDocumentBuilder.InsertHtml.docx");
            //ExEnd:InsertHtml
        }

        [Test]
        public void InsertHyperlink()
        {
            //ExStart:InsertHyperlink
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Write("Please make sure to visit ");
            builder.Font.Color = Color.Blue;
            builder.Font.Underline = Underline.Single;
            
            builder.InsertHyperlink("Aspose Website", "http://www.aspose.com", false);
            
            builder.Font.ClearFormatting();
            builder.Write(" for more information.");

            doc.Save(ArtifactsDir + "AddContentUsingDocumentBuilder.InsertHyperlink.docx");
            //ExEnd:InsertHyperlink
        }

        [Test]
        public void InsertTableOfContents()
        {
            //ExStart:InsertTableOfContents
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            
            // Start the actual document content on the second page.
            builder.InsertBreak(BreakType.PageBreak);

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

            //ExStart:UpdateFields
            // The newly inserted table of contents will be initially empty.
            // It needs to be populated by updating the fields in the document.
            doc.UpdateFields();
            //ExEnd:UpdateFields

            doc.Save(ArtifactsDir + "AddContentUsingDocumentBuilder.InsertTableOfContents.docx");
            //ExEnd:InsertTableOfContents
        }

        [Test]
        public void InsertInlineImage()
        {
            //ExStart:InsertInlineImage
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage(ImagesDir + "Transparent background logo.png");

            doc.Save(ArtifactsDir + "AddContentUsingDocumentBuilder.InsertInlineImage.docx");
            //ExEnd:InsertInlineImage
        }

        [Test]
        public void InsertFloatingImage()
        {
            //ExStart:InsertFloatingImage
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage(ImagesDir + "Transparent background logo.png",
                RelativeHorizontalPosition.Margin,
                100,
                RelativeVerticalPosition.Margin,
                100,
                200,
                100,
                WrapType.Square);

            doc.Save(ArtifactsDir + "AddContentUsingDocumentBuilder.InsertFloatingImage.docx");
            //ExEnd:InsertFloatingImage
        }

        [Test]
        public void InsertParagraph()
        {
            //ExStart:InsertParagraph
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Font font = builder.Font;
            font.Size = 16;
            font.Bold = true;
            font.Color = Color.Blue;
            font.Name = "Arial";
            font.Underline = Underline.Dash;

            ParagraphFormat paragraphFormat = builder.ParagraphFormat;
            paragraphFormat.FirstLineIndent = 8;
            paragraphFormat.Alignment = ParagraphAlignment.Justify;
            paragraphFormat.KeepTogether = true;

            builder.Writeln("A whole paragraph.");

            doc.Save(ArtifactsDir + "AddContentUsingDocumentBuilder.InsertParagraph.docx");
            //ExEnd:InsertParagraph
        }

        [Test]
        public void InsertTCField()
        {
            //ExStart:InsertTCField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField("TC \"Entry Text\" \\f t");

            doc.Save(ArtifactsDir + "AddContentUsingDocumentBuilder.InsertTCField.docx");
            //ExEnd:InsertTCField
        }

        [Test]
        public void InsertTCFieldsAtText()
        {
            //ExStart:InsertTCFieldsAtText
            Document doc = new Document();

            FindReplaceOptions options = new FindReplaceOptions();
            options.ApplyFont.HighlightColor = Color.DarkOrange;
            options.ReplacingCallback = new InsertTCFieldHandler("Chapter 1", "\\l 1");

            doc.Range.Replace(new Regex("The Beginning"), "", options);
            //ExEnd:InsertTCFieldsAtText
        }

        //ExStart:InsertTCFieldHandler
        public sealed class InsertTCFieldHandler : IReplacingCallback
        {
            // Store the text and switches to be used for the TC fields.
            private readonly string mFieldText;
            private readonly string mFieldSwitches;

            /// <summary>
            /// The display text and switches to use for each TC field. Display name can be an empty string or null.
            /// </summary>
            public InsertTCFieldHandler(string text, string switches)
            {
                mFieldText = text;
                mFieldSwitches = switches;
            }

            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
            {
                DocumentBuilder builder = new DocumentBuilder((Document) args.MatchNode.Document);
                builder.MoveTo(args.MatchNode);

                // If the user-specified text to be used in the field as display text, then use that,
                // otherwise use the match string as the display text.
                string insertText = !string.IsNullOrEmpty(mFieldText) ? mFieldText : args.Match.Value;

                builder.InsertField($"TC \"{insertText}\" {mFieldSwitches}");

                return ReplaceAction.Skip;
            }
        }
        //ExEnd:InsertTCFieldHandler
        
        [Test]
        public void CursorPosition()
        {
            //ExStart:CursorPosition
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Node curNode = builder.CurrentNode;
            Paragraph curParagraph = builder.CurrentParagraph;
            //ExEnd:CursorPosition

            Console.WriteLine("\nCursor move to paragraph: " + curParagraph.GetText());
        }

        [Test]
        public void MoveToNode()
        {
            //ExStart:MoveToNode
            //ExStart:MoveToBookmark
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a bookmark and add content to it using a DocumentBuilder.
            builder.StartBookmark("MyBookmark");
            builder.Writeln("Bookmark contents.");
            builder.EndBookmark("MyBookmark");

            // The node that the DocumentBuilder is currently at is past the boundaries of the bookmark.
            Assert.AreEqual(doc.Range.Bookmarks[0].BookmarkEnd, builder.CurrentParagraph.FirstChild);

            // If we wish to revise the content of our bookmark with the DocumentBuilder, we can move back to it like this.
            builder.MoveToBookmark("MyBookmark");

            // Now we're located between the bookmark's BookmarkStart and BookmarkEnd nodes, so any text the builder adds will be within it.
            Assert.AreEqual(doc.Range.Bookmarks[0].BookmarkStart, builder.CurrentParagraph.FirstChild);

            // We can move the builder to an individual node,
            // which in this case will be the first node of the first paragraph, like this.
            builder.MoveTo(doc.FirstSection.Body.FirstParagraph.GetChildNodes(NodeType.Any, false)[0]);
            //ExEnd:MoveToBookmark

            Assert.AreEqual(NodeType.BookmarkStart, builder.CurrentNode.NodeType);
            Assert.IsTrue(builder.IsAtStartOfParagraph);

            // A shorter way of moving the very start/end of a document is with these methods.
            builder.MoveToDocumentEnd();
            Assert.IsTrue(builder.IsAtEndOfParagraph);
            builder.MoveToDocumentStart();
            Assert.IsTrue(builder.IsAtStartOfParagraph);
            //ExEnd:MoveToNode
        }

        [Test]
        public void MoveToDocumentStartEnd()
        {
            //ExStart:MoveToDocumentStartEnd
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the cursor position to the beginning of your document.
            builder.MoveToDocumentStart();
            Console.WriteLine("\nThis is the beginning of the document.");

            // Move the cursor position to the end of your document.
            builder.MoveToDocumentEnd();
            Console.WriteLine("\nThis is the end of the document.");
            //ExEnd:MoveToDocumentStartEnd            
        }

        [Test]
        public void MoveToSection()
        {
            //ExStart:MoveToSection
            Document doc = new Document();
            doc.AppendChild(new Section(doc));

            // Move a DocumentBuilder to the second section and add text.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToSection(1);
            builder.Writeln("Text added to the 2nd section.");

            // Create document with paragraphs.
            doc = new Document(MyDir + "Paragraphs.docx");
            ParagraphCollection paragraphs = doc.FirstSection.Body.Paragraphs;
            Assert.AreEqual(23, paragraphs.Count);

            // When we create a DocumentBuilder for a document, its cursor is at the very beginning of the document by default,
            // and any content added by the DocumentBuilder will just be prepended to the document.
            builder = new DocumentBuilder(doc);
            Assert.AreEqual(0, paragraphs.IndexOf(builder.CurrentParagraph));

            // You can move the cursor to any position in a paragraph.
            builder.MoveToParagraph(2, 14);
            Assert.AreEqual(2, paragraphs.IndexOf(builder.CurrentParagraph));
            builder.Writeln("This is a new third paragraph. ");
            Assert.AreEqual(3, paragraphs.IndexOf(builder.CurrentParagraph));
            //ExEnd:MoveToSection               
        }

        [Test]
        public void MoveToHeadersFooters()
        {
            //ExStart:MoveToHeadersFooters
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Specify that we want headers and footers different for first, even and odd pages.
            builder.PageSetup.DifferentFirstPageHeaderFooter = true;
            builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

            // Create the headers.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.Write("Header for the first page");
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
            builder.Write("Header for even pages");
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Header for all other pages");

            // Create two pages in the document.
            builder.MoveToSection(0);
            builder.Writeln("Page1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page2");

            doc.Save(ArtifactsDir + "AddContentUsingDocumentBuilder.MoveToHeadersFooters.docx");
            //ExEnd:MoveToHeadersFooters
        }

        [Test]
        public void MoveToParagraph()
        {
            //ExStart:MoveToParagraph
            Document doc = new Document(MyDir + "Paragraphs.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToParagraph(2, 0);
            builder.Writeln("This is the 3rd paragraph.");
            //ExEnd:MoveToParagraph               
        }

        [Test]
        public void MoveToTableCell()
        {
            //ExStart:MoveToTableCell
            Document doc = new Document(MyDir + "Tables.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the builder to row 3, cell 4 of the first table.
            builder.MoveToCell(0, 2, 3, 0);
            builder.Write("\nCell contents added by DocumentBuilder");
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            Assert.AreEqual(table.Rows[2].Cells[3], builder.CurrentNode.ParentNode.ParentNode);
            Assert.AreEqual("Cell contents added by DocumentBuilderCell 3 contents\a", table.Rows[2].Cells[3].GetText().Trim());
            //ExEnd:MoveToTableCell               
        }

        [Test]
        public void MoveToBookmarkEnd()
        {
            //ExStart:MoveToBookmarkEnd
            Document doc = new Document(MyDir + "Bookmarks.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToBookmark("MyBookmark1", false, true);
            builder.Writeln("This is a bookmark.");
            //ExEnd:MoveToBookmarkEnd              
        }

        [Test]
        public void MoveToMergeField()
        {
            //ExStart:MoveToMergeField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a field using the DocumentBuilder and add a run of text after it.
            Field field = builder.InsertField("MERGEFIELD field");
            builder.Write(" Text after the field.");

            // The builder's cursor is currently at end of the document.
            Assert.Null(builder.CurrentNode);
            // We can move the builder to a field like this, placing the cursor at immediately after the field.
            builder.MoveToField(field, true);

            // Note that the cursor is at a place past the FieldEnd node of the field, meaning that we are not actually inside the field.
            // If we wish to move the DocumentBuilder to inside a field,
            // we will need to move it to a field's FieldStart or FieldSeparator node using the DocumentBuilder.MoveTo() method.
            Assert.AreEqual(field.End, builder.CurrentNode.PreviousSibling);
            builder.Write(" Text immediately after the field.");
            //ExEnd:MoveToMergeField              
        }        
    }
}