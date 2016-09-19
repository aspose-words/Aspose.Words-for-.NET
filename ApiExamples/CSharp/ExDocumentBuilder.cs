// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;
using System.IO;

using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Fields;
using Aspose.Words.Tables;


using NUnit.Framework;

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
            //ExId:DocumentBuilderInsertText
            //ExSummary:Inserts formatted text using DocumentBuilder.
            DocumentBuilder builder = new DocumentBuilder();

            // Specify font formatting before adding text.
            Aspose.Words.Font font = builder.Font;
            font.Size = 16;
            font.Bold = true;
            font.Color = Color.Blue;
            font.Name = "Arial";
            font.Underline = Underline.Dash;

            builder.Write("Sample text.");
            //ExEnd
        }

        [Test]
        public void HeadersAndFooters()
        {
            //ExStart
            //ExFor:DocumentBuilder.#ctor(Document)
            //ExFor:DocumentBuilder.MoveToHeaderFooter
            //ExFor:DocumentBuilder.MoveToSection
            //ExFor:DocumentBuilder.InsertBreak
            //ExFor:HeaderFooterType
            //ExFor:PageSetup.DifferentFirstPageHeaderFooter
            //ExFor:PageSetup.OddAndEvenPagesHeaderFooter
            //ExFor:BreakType
            //ExId:DocumentBuilderMoveToHeaderFooter
            //ExSummary:Creates headers and footers in a document using DocumentBuilder.
            // Create a blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Specify that we want headers and footers different for first, even and odd pages.
            builder.PageSetup.DifferentFirstPageHeaderFooter = true;
            builder.PageSetup.OddAndEvenPagesHeaderFooter = true;

            // Create the headers.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderFirst);
            builder.Write("Header First");
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderEven);
            builder.Write("Header Even");
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Header Odd");

            // Create three pages in the document.
            builder.MoveToSection(0);
            builder.Writeln("Page1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page3");

            doc.Save(MyDir + @"\Artifacts\DocumentBuilder.HeadersAndFooters.doc");
            //ExEnd
        }

        [Test]
        public void InsertMergeField()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertField(string)
            //ExId:DocumentBuilderInsertField
            //ExSummary:Inserts a merge field into a document using DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField(@"MERGEFIELD MyFieldName \* MERGEFORMAT");
            //ExEnd			
        }

        [Test]
        public void InsertField()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertField(string)
            //ExFor:Field
            //ExFor:Field.Update
            //ExFor:Field.Result
            //ExFor:Field.GetFieldCode
            //ExFor:Field.Type
            //ExFor:Field.Remove
            //ExFor:FieldType
            //ExSummary:Inserts a field into a document using DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a simple Date field into the document.
            // When we insert a field through the DocumentBuilder class we can get the
            // special Field object which contains information about the field.
            Field dateField = builder.InsertField(@"DATE \* MERGEFORMAT");

            // Update this particular field in the document so we can get the FieldResult.
            dateField.Update();

            // Display some information from this field.
            // The field result is where the last evaluated value is stored. This is what is displayed in the document
            // When field codes are not showing.
            Console.WriteLine("FieldResult: {0}", dateField.Result);

            // Display the field code which defines the behaviour of the field. This can been seen in Microsoft Word by pressing ALT+F9.
            Console.WriteLine("FieldCode: {0}", dateField.GetFieldCode());

            // The field type defines what type of field in the Document this is. In this case the type is "FieldDate" 
            Console.WriteLine("FieldType: {0}", dateField.Type);

            // Finally let's completely remove the field from the document. This can easily be done by invoking the Remove method on the object.
            dateField.Remove();
            //ExEnd			
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void GetFieldCode(bool nestedFields)
        {
            Document doc = new Document(MyDir + "Field.FieldCode.docx");

            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type == FieldType.FieldIf)
                {
                    FieldIf fif = (FieldIf)field;

                    Assert.AreEqual(" IF  MERGEFIELD Q223  > 0 \" (and additionally London Weighting of   MERGEFIELD  Q223 \\f £  per hour) \" \"\" ", fif.GetFieldCode());

                    if (nestedFields)
                    {
                        Assert.AreEqual(" IF  MERGEFIELD Q223  > 0 \" (and additionally London Weighting of   MERGEFIELD  Q223 \\f £  per hour) \" \"\" ", fif.GetFieldCode(true));
                    }
                    else
                    {
                        Assert.AreEqual(" IF  > 0 \" (and additionally London Weighting of   per hour) \" \"\" ", fif.GetFieldCode(false));
                    }
                }
            }
        }
        
        [Test]
        public void DocumentBuilderAndSave()
        {
            //ExStart
            //ExId:DocumentBuilderAndSave
            //ExSummary:Shows how to create build a document using a document builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello World!");

            doc.Save(MyDir + @"\Artifacts\DocumentBuilderAndSave.docx");
            //ExEnd
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
            //ExId:DocumentBuilderInsertHyperlink
            //ExSummary:Inserts a hyperlink into a document using DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Please make sure to visit ");

            // Specify font formatting for the hyperlink.
            builder.Font.Color = Color.Blue;
            builder.Font.Underline = Underline.Single;
            // Insert the link.
            builder.InsertHyperlink("Aspose Website", "http://www.aspose.com", false);

            // Revert to default formatting.
            builder.Font.ClearFormatting();

            builder.Write(" for more information.");

            doc.Save(MyDir + @"\Artifacts\DocumentBuilder.InsertHyperlink.doc");
            //ExEnd
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

            // Set up font formatting and write text that goes before the hyperlink.
            builder.Font.Name = "Arial";
            builder.Font.Size = 24;
            builder.Font.Bold = true;
            builder.Write("To go to an important location, click ");

            // Save the font formatting so we use different formatting for hyperlink and restore old formatting later.
            builder.PushFont();

            // Set new font formatting for the hyperlink and insert the hyperlink.
            // The "Hyperlink" style is a Microsoft Word built-in style so we don't have to worry to 
            // create it, it will be created automatically if it does not yet exist in the document.
            builder.Font.StyleIdentifier = StyleIdentifier.Hyperlink;
            builder.InsertHyperlink("here", "http://www.google.com", false);

            // Restore the formatting that was before the hyperlink.
            builder.PopFont();

            builder.Writeln(". We hope you enjoyed the example.");

            doc.Save(MyDir + @"\Artifacts\DocumentBuilder.PushPopFont.doc");
            //ExEnd
        }

        [Test]
        public void InsertWatermark()
        {
            //ExStart
            //ExFor:HeaderFooterType
            //ExFor:DocumentBuilder.MoveToHeaderFooter
            //ExFor:PageSetup.PageWidth
            //ExFor:PageSetup.PageHeight
            //ExFor:DocumentBuilder.InsertImage(Image)
            //ExFor:WrapType
            //ExFor:RelativeHorizontalPosition
            //ExFor:RelativeVerticalPosition
            //ExSummary:Inserts a watermark image into a document using DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The best place for the watermark image is in the header or footer so it is shown on every page.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            Image image = Image.FromFile(MyDir + "Watermark.png");

            // Insert a floating picture.
            Shape shape = builder.InsertImage(image);
            shape.WrapType = WrapType.None;
            shape.BehindText = true;

            shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;

            // Calculate image left and top position so it appears in the centre of the page.
            shape.Left = (builder.PageSetup.PageWidth - shape.Width) / 2;
            shape.Top = (builder.PageSetup.PageHeight - shape.Height) / 2;

            doc.Save(MyDir + @"\Artifacts\DocumentBuilder.InsertWatermark.doc");
            //ExEnd
        }

        [Test]
        public void InsertHtml()
        {
            //ExStart
            //ExFor:DocumentBuilder
            //ExFor:DocumentBuilder.InsertHtml(string)
            //ExId:DocumentBuilderInsertHtml
            //ExSummary:Inserts HTML into a document. The formatting specified in the HTML is applied.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertHtml(
                "<P align='right'>Paragraph right</P>" +
                "<b>Implicit paragraph left</b>" +
                "<div align='center'>Div center</div>" +
                "<h1 align='left'>Heading 1 left.</h1>");

            doc.Save(MyDir + @"\Artifacts\DocumentBuilder.InsertHtml.doc");
            //ExEnd
        }

        [Test]
        public void InsertHtmlEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertHtml(String, Boolean)
            //ExSummary:Inserts HTML into a document using. The current document formatting at the insertion position is applied to the inserted text. 
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            bool useBuilderFormatting = true;

            builder.InsertHtml(
                "<P align='right'>Paragraph right</P>" +
                "<b>Implicit paragraph left</b>" +
                "<div align='center'>Div center</div>" +
                "<h1 align='left'>Heading 1 left.</h1>", useBuilderFormatting);

            doc.Save(MyDir + @"\Artifacts\DocumentBuilder.InsertHtml.doc");
            //ExEnd
        }

        //For assert this test you need to open "MathML.docx" and "MathML.pdf" and check, that mathml code are render as "a 1 + b 1"
        [Test]
        public void InsertMathMl()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            const string MathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><msub><mi>a</mi><mrow><mn>1</mn></mrow></msub><mo>+</mo><msub><mi>b</mi><mrow><mn>1</mn></mrow></msub></mrow></math>";

            builder.InsertHtml(MathMl);

            doc.Save(MyDir + "MathML.docx");
            doc.Save(MyDir + "MathML.pdf");
        }

        [Test]
        public void InsertTextAndBookmark()
        {
            //ExStart
            //ExFor:DocumentBuilder
            //ExFor:DocumentBuilder.StartBookmark
            //ExFor:DocumentBuilder.EndBookmark
            //ExSummary:Adds some text into the document and encloses the text in a bookmark using DocumentBuilder.
            DocumentBuilder builder = new DocumentBuilder();

            builder.StartBookmark("MyBookmark");
            builder.Writeln("Text inside a bookmark.");
            builder.EndBookmark("MyBookmark");
            //ExEnd
        }

        [Test]
        public void CreateForm()
        {
            //ExStart
            //ExFor:TextFormFieldType
            //ExFor:DocumentBuilder.InsertTextInput
            //ExFor:DocumentBuilder.InsertComboBox
            //ExSummary:Builds a sample form to fill.
            DocumentBuilder builder = new DocumentBuilder();

            // Insert a text form field for input a name.
            builder.InsertTextInput("", TextFormFieldType.Regular, "", "Enter your name here", 30);

            // Insert two blank lines.
            builder.Writeln("");
            builder.Writeln("");

            string[] items = new string[]
                {
                    "-- Select your favorite footwear --",
                    "Sneakers",
                    "Oxfords",
                    "Flip-flops",
                    "Other",
                    "I prefer to be barefoot"
                };

            // Insert a combo box to select a footwear type.
            builder.InsertComboBox("", items, 0);

            // Insert 2 blank lines.
            builder.Writeln("");
            builder.Writeln("");

            builder.Document.Save(MyDir + @"\Artifacts\DocumentBuilder.CreateForm.doc");
            //ExEnd
        }

        [Ignore("Bug \"trimmed name if you enter more than 20 characters\"")]
        [Test]
        public void InsertCheckBox()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertCheckBox(string, bool, bool, int)
            //ExFor:DocumentBuilder.InsertCheckBox(string, bool, int)
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            //Insert checkboxes
            //With Default value
            builder.InsertCheckBox("CheckBox_DefaultAndCheckedValue", false, true, 0);

            //Without Default value
            builder.InsertCheckBox("CheckBox_OnlyCheckedValue", true, 100);
            //ExEnd

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            //Get checkboxes from the document
            FormFieldCollection formFields = doc.Range.FormFields;

            //Check that is the right checkbox
            Assert.AreEqual("CheckBox_DefaultAndCheckedValue", formFields[0].Name);

            //Assert that parameters sets correctly
            Assert.AreEqual(true, formFields[0].Checked);
            Assert.AreEqual(false, formFields[0].Default);
            Assert.AreEqual(10, formFields[0].CheckBoxSize);

            //Check that is the right checkbox
            Assert.AreEqual("CheckBox_OnlyCheckedValue", formFields[1].Name);

            //Assert that parameters sets correctly
            Assert.AreEqual(false, formFields[1].Checked);
            Assert.AreEqual(false, formFields[1].Default);
            Assert.AreEqual(100, formFields[1].CheckBoxSize);
        }

        [Test]
        public void InsertCheckBoxEmptyName()
        {
            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);

            //Assert that empty string name working correctly
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
            //ExSummary:Shows how to move between nodes and manipulate current ones.
            Document doc = new Document(MyDir + "DocumentBuilder.WorkingWithNodes.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move to a bookmark and delete the parent paragraph.
            builder.MoveToBookmark("ParaToDelete");
            builder.CurrentParagraph.Remove();

            // Move to a particular paragraph's run and replace all occurrences of "bad" with "good" within this run.
            builder.MoveTo(doc.LastSection.Body.Paragraphs[0].Runs[0]);
            builder.CurrentNode.Range.Replace("bad", "good", false, true);

            // Mark the beginning of the document.
            builder.MoveToDocumentStart();
            builder.Writeln("Start of document.");

            // Mark the ending of the document.
            builder.MoveToDocumentEnd();
            builder.Writeln("End of document.");

            doc.Save(MyDir + @"\Artifacts\DocumentBuilder.WorkingWithNodes.doc");
            //ExEnd
        }

        [Test]
        public void FillingDocument()
        {
            //ExStart
            //ExFor:DocumentBuilder.MoveToMergeField(string)
            //ExFor:DocumentBuilder.Bold
            //ExFor:DocumentBuilder.Italic
            //ExSummary:Fills document merge fields with some data.
            Document doc = new Document(MyDir + "DocumentBuilder.FillingDocument.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToMergeField("TeamLeaderName");
            builder.Bold = true;
            builder.Writeln("Roman Korchagin");

            builder.MoveToMergeField("SoftwareDeveloper1Name");
            builder.Italic = true;
            builder.Writeln("Dmitry Vorobyev");

            builder.MoveToMergeField("SoftwareDeveloper2Name");
            builder.Italic = true;
            builder.Writeln("Vladimir Averkin");

            doc.Save(MyDir + @"\Artifacts\DocumentBuilder.FillingDocument.doc");
            //ExEnd
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
            //ExId:InsertTableOfContents
            //ExSummary:Demonstrates how to insert a Table of contents (TOC) into a document using heading styles as entries.
            // Use a blank document
            Document doc = new Document();

            // Create a document builder to insert content with into document.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table of contents at the beginning of the document.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

            // Start the actual document content on the second page.
            builder.InsertBreak(BreakType.PageBreak);

            // Build a document with complex structure by applying different heading styles thus creating TOC entries.
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

            // Call the method below to update the TOC.
            doc.UpdateFields();
            //ExEnd

            doc.Save(MyDir + @"\Artifacts\DocumentBuilder.InsertToc.docx");
        }

        [Test]
        public void InsertTable()
        {
            //ExStart
            //ExFor:DocumentBuilder
            //ExFor:DocumentBuilder.StartTable
            //ExFor:DocumentBuilder.InsertCell
            //ExFor:DocumentBuilder.EndRow
            //ExFor:DocumentBuilder.EndTable
            //ExFor:DocumentBuilder.CellFormat
            //ExFor:DocumentBuilder.RowFormat
            //ExFor:CellFormat
            //ExFor:CellFormat.Width
            //ExFor:CellFormat.VerticalAlignment
            //ExFor:CellFormat.Shading
            //ExFor.CellFormat.Orientation
            //ExFor:RowFormat
            //ExFor:RowFormat.HeightRule
            //ExFor:RowFormat.Height
            //ExFor:RowFormat.Borders
            //ExFor:Shading.BackgroundPatternColor
            //ExFor:Shading.ClearFormatting
            //ExSummary:Shows how to build a nice bordered table.
            DocumentBuilder builder = new DocumentBuilder();

            // Start building a table.
            builder.StartTable();

            // Set the appropriate paragraph, cell, and row formatting. The formatting properties are preserved
            // until they are explicitly modified so there's no need to set them for each row or cell. 

            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            builder.CellFormat.Width = 300;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.GreenYellow;

            builder.RowFormat.HeightRule = HeightRule.Exactly;
            builder.RowFormat.Height = 50;
            builder.RowFormat.Borders.LineStyle = LineStyle.Engrave3D;
            builder.RowFormat.Borders.Color = Color.Orange;

            builder.InsertCell();
            builder.Write("Row 1, Col 1");

            builder.InsertCell();
            builder.Write("Row 1, Col 2");

            builder.EndRow();

            // Remove the shading (clear background).
            builder.CellFormat.Shading.ClearFormatting();

            builder.InsertCell();
            builder.Write("Row 2, Col 1");

            builder.InsertCell();
            builder.Write("Row 2, Col 2");

            builder.EndRow();

            builder.InsertCell();

            // Make the row height bigger so that a vertically oriented text could fit into cells.
            builder.RowFormat.Height = 150;
            builder.CellFormat.Orientation = TextOrientation.Upward;
            builder.Write("Row 3, Col 1");

            builder.InsertCell();
            builder.CellFormat.Orientation = TextOrientation.Downward;
            builder.Write("Row 3, Col 2");

            builder.EndRow();

            builder.EndTable();

            builder.Document.Save(MyDir + @"\Artifacts\DocumentBuilder.InsertTable.doc");
            //ExEnd
        }

        [Test]
        public void InsertTableWithTableStyle()
        {
            //ExStart
            //ExFor:Table.StyleIdentifier
            //ExFor:Table.StyleOptions
            //ExFor:TableStyleOptions
            //ExFor:Table.AutoFit
            //ExFor:AutoFitBehavior
            //ExId:InsertTableWithTableStyle
            //ExSummary:Shows how to build a new table with a table style applied.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            // We must insert at least one row first before setting any table formatting.
            builder.InsertCell();
            // Set the table style used based of the unique style identifier.
            // Note that not all table styles are available when saving as .doc format.
            table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;
            // Apply which features should be formatted by the style.
            table.StyleOptions = TableStyleOptions.FirstColumn | TableStyleOptions.RowBands | TableStyleOptions.FirstRow;
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Continue with building the table as normal.
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

            doc.Save(MyDir + @"\Artifacts\DocumentBuilder.SetTableStyle.docx");
            //ExEnd

            // Verify that the style was set by expanding to direct formatting.
            doc.ExpandTableStylesToDirectFormatting();
            Assert.AreEqual("Medium Shading 1 Accent 1", table.Style.Name);
            Assert.AreEqual(TableStyleOptions.FirstColumn | TableStyleOptions.RowBands | TableStyleOptions.FirstRow, table.StyleOptions);
            Assert.AreEqual(189, table.FirstRow.FirstCell.CellFormat.Shading.BackgroundPatternColor.B);
            Assert.AreEqual(Color.White.ToArgb(), table.FirstRow.FirstCell.FirstParagraph.Runs[0].Font.Color.ToArgb());
            Assert.AreNotEqual(Color.LightBlue.ToArgb(), table.LastRow.FirstCell.CellFormat.Shading.BackgroundPatternColor.B);
            Assert.AreEqual(Color.Empty.ToArgb(), table.LastRow.FirstCell.FirstParagraph.Runs[0].Font.Color.ToArgb());
        }

        [Test]
        public void InsertTableSetHeadingRow()
        {
            //ExStart
            //ExFor:RowFormat.HeadingFormat
            //ExId:InsertTableWithHeadingFormat
            //ExSummary:Shows how to build a table which include heading rows that repeat on subsequent pages. 
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
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

            // Insert some content so the table is long enough to continue onto the next page.
            for (int i = 0; i < 50; i++)
            {
                builder.InsertCell();
                builder.RowFormat.HeadingFormat = false;
                builder.Write("Column 1 Text");
                builder.InsertCell();
                builder.Write("Column 2 Text");
                builder.EndRow();
            }

            doc.Save(MyDir + @"\Artifacts\Table.HeadingRow.doc");
            //ExEnd

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
            //ExId:TablePreferredWidth
            //ExSummary:Shows how to set a table to auto fit to 50% of the page width.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table with a width that takes up half the page width.
            Table table = builder.StartTable();

            // Insert a few cells
            builder.InsertCell();
            table.PreferredWidth = PreferredWidth.FromPercent(50);
            builder.Writeln("Cell #1");

            builder.InsertCell();
            builder.Writeln("Cell #2");

            builder.InsertCell();
            builder.Writeln("Cell #3");

            doc.Save(MyDir + @"\Artifacts\Table.PreferredWidth.doc");
            //ExEnd

            // Verify the correct settings were applied.
            Assert.AreEqual(PreferredWidthType.Percent, table.PreferredWidth.Type);
            Assert.AreEqual(50, table.PreferredWidth.Value);
        }

        [Test]
        public void InsertCellsWithDifferentPreferredCellWidths()
        {
            //ExStart
            //ExFor:CellFormat.PreferredWidth
            //ExFor:PreferredWidth
            //ExFor:PreferredWidth.FromPoints
            //ExFor:PreferredWidth.FromPercent
            //ExFor:PreferredWidth.Auto
            //ExId:CellPreferredWidths
            //ExSummary:Shows how to set the different preferred width settings.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table row made up of three cells which have different preferred widths.
            Table table = builder.StartTable();

            // Insert an absolute sized cell.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(40);
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightYellow;
            builder.Writeln("Cell at 40 points width");

            // Insert a relative (percent) sized cell.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPercent(20);
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            builder.Writeln("Cell at 20% width");

            // Insert a auto sized cell.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.Auto;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
            builder.Writeln("Cell automatically sized. The size of this cell is calculated from the table preferred width.");
            builder.Writeln("In this case the cell will fill up the rest of the available space.");

            doc.Save(MyDir + @"\Artifacts\Table.CellPreferredWidths.doc");
            //ExEnd

            // Verify the correct settings were applied.
            Assert.AreEqual(PreferredWidthType.Points, table.FirstRow.FirstCell.CellFormat.PreferredWidth.Type);
            Assert.AreEqual(PreferredWidthType.Percent, table.FirstRow.Cells[1].CellFormat.PreferredWidth.Type);
            Assert.AreEqual(PreferredWidthType.Auto, table.FirstRow.Cells[2].CellFormat.PreferredWidth.Type);
        }

        [Test]
        public void InsertTableFromHtml()
        {
            //ExStart
            //ExId:InsertTableFromHtml
            //ExSummary:Shows how to insert a table in a document from a string containing HTML tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the table from HTML. Note that AutoFitSettings does not apply to tables
            // inserted from HTML.
            builder.InsertHtml("<table>" +
                               "<tr>" +
                               "<td>Row 1, Cell 1</td>" +
                               "<td>Row 1, Cell 2</td>" +
                               "</tr>" +
                               "<tr>" +
                               "<td>Row 2, Cell 2</td>" +
                               "<td>Row 2, Cell 2</td>" +
                               "</tr>" +
                               "</table>");

            doc.Save(MyDir + @"\Artifacts\DocumentBuilder.InsertTableFromHtml.doc");
            //ExEnd

            // Verify the table was constructed properly.
            Assert.AreEqual(1, doc.GetChildNodes(NodeType.Table, true).Count);
            Assert.AreEqual(2, doc.GetChildNodes(NodeType.Row, true).Count);
            Assert.AreEqual(4, doc.GetChildNodes(NodeType.Cell, true).Count);
        }

        [Test]
        public void BuildNestedTableUsingDocumentBuilder()
        {
            //ExStart
            //ExFor:Cell.FirstParagraph
            //ExId:BuildNestedTableUsingDocumentBuilder
            //ExSummary:Shows how to insert a nested table using DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build the outer table.
            Cell cell = builder.InsertCell();
            builder.Writeln("Outer Table Cell 1");

            builder.InsertCell();
            builder.Writeln("Outer Table Cell 2");

            // This call is important in order to create a nested table within the first table
            // Without this call the cells inserted below will be appended to the outer table.
            builder.EndTable();

            // Move to the first cell of the outer table.
            builder.MoveTo(cell.FirstParagraph);

            // Build the inner table.
            builder.InsertCell();
            builder.Writeln("Inner Table Cell 1");
            builder.InsertCell();
            builder.Writeln("Inner Table Cell 2");

            builder.EndTable();

            doc.Save(MyDir + @"\Artifacts\DocumentBuilder.InsertNestedTable.doc");
            //ExEnd

            Assert.AreEqual(2, doc.GetChildNodes(NodeType.Table, true).Count);
            Assert.AreEqual(4, doc.GetChildNodes(NodeType.Cell, true).Count);
            Assert.AreEqual(1, cell.Tables[0].Count);
            Assert.AreEqual(2, cell.Tables[0].FirstRow.Cells.Count);
        }

        [Test]
        public void BuildSimpleTable()
        {
            //ExStart
            //ExFor:DocumentBuilder
            //ExFor:DocumentBuilder.Write
            //ExFor:DocumentBuilder.InsertCell
            //ExId:BuildSimpleTable
            //ExSummary:Shows how to create a simple table using DocumentBuilder with default formatting.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // We call this method to start building the table.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Row 1, Cell 1 Content.");

            // Build the second cell
            builder.InsertCell();
            builder.Write("Row 1, Cell 2 Content.");
            // Call the following method to end the row and start a new row.
            builder.EndRow();

            // Build the first cell of the second row.
            builder.InsertCell();
            builder.Write("Row 2, Cell 1 Content");

            // Build the second cell.
            builder.InsertCell();
            builder.Write("Row 2, Cell 2 Content.");
            builder.EndRow();

            // Signal that we have finished building the table.
            builder.EndTable();

            // Save the document to disk.
            doc.Save(MyDir + @"\Artifacts\DocumentBuilder.CreateSimpleTable.doc");
            //ExEnd

            // Verify that the cell count of the table is four.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            Assert.IsNotNull(table);
            Assert.AreEqual(table.GetChildNodes(NodeType.Cell, true).Count, 4);
        }

        [Test]
        public void BuildFormattedTable()
        {
            //ExStart
            //ExFor:DocumentBuilder
            //ExFor:DocumentBuilder.Write
            //ExFor:DocumentBuilder.InsertCell
            //ExFor:RowFormat.Height
            //ExFor:RowFormat.HeightRule
            //ExFor:Table.LeftIndent
            //ExFor:Shading.BackgroundPatternColor
            //ExFor:DocumentBuilder.ParagraphFormat
            //ExFor:DocumentBuilder.Font
            //ExId:BuildFormattedTable
            //ExSummary:Shows how to create a formatted table using DocumentBuilder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();

            // Make the header row.
            builder.InsertCell();

            // Set the left indent for the table. Table wide formatting must be applied after 
            // at least one row is present in the table.
            table.LeftIndent = 20.0;

            // Set height and define the height rule for the header row.
            builder.RowFormat.Height = 40.0;
            builder.RowFormat.HeightRule = HeightRule.AtLeast;

            // Some special features for the header row.
            builder.CellFormat.Shading.BackgroundPatternColor = Color.FromArgb(198, 217, 241);
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.Font.Size = 16;
            builder.Font.Name = "Arial";
            builder.Font.Bold = true;

            builder.CellFormat.Width = 100.0;
            builder.Write("Header Row,\n Cell 1");

            // We don't need to specify the width of this cell because it's inherited from the previous cell.
            builder.InsertCell();
            builder.Write("Header Row,\n Cell 2");

            builder.InsertCell();
            builder.CellFormat.Width = 200.0;
            builder.Write("Header Row,\n Cell 3");
            builder.EndRow();

            // Set features for the other rows and cells.
            builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
            builder.CellFormat.Width = 100.0;
            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;

            // Reset height and define a different height rule for table body
            builder.RowFormat.Height = 30.0;
            builder.RowFormat.HeightRule = HeightRule.Auto;
            builder.InsertCell();
            // Reset font formatting.
            builder.Font.Size = 12;
            builder.Font.Bold = false;

            // Build the other cells.
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

            doc.Save(MyDir + @"\Artifacts\DocumentBuilder.CreateFormattedTable.doc");
            //ExEnd

            // Verify that the cell style is different compared to default.
            Assert.AreNotEqual(table.LeftIndent, 0.0);
            Assert.AreNotEqual(table.FirstRow.RowFormat.HeightRule, HeightRule.Auto);
            Assert.AreNotEqual(table.FirstRow.FirstCell.CellFormat.Shading.BackgroundPatternColor, Color.Empty);
            Assert.AreNotEqual(table.FirstRow.FirstCell.FirstParagraph.ParagraphFormat.Alignment, ParagraphAlignment.Left);
        }

        [Test]
        public void SetCellShadingAndBorders()
        {
            //ExStart
            //ExFor:Shading
            //ExFor:Shading.BackgroundPatternColor
            //ExFor:Table.SetBorders
            //ExFor:BorderCollection.Left
            //ExFor:BorderCollection.Right
            //ExFor:BorderCollection.Top
            //ExFor:BorderCollection.Bottom
            //ExId:TableBordersAndShading
            //ExSummary:Shows how to format table and cell with different borders and shadings
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();

            // Set the borders for the entire table.
            table.SetBorders(LineStyle.Single, 2.0, Color.Black);
            // Set the cell shading for this cell.
            builder.CellFormat.Shading.BackgroundPatternColor = Color.Red;
            builder.Writeln("Cell #1");

            builder.InsertCell();
            // Specify a different cell shading for the second cell.
            builder.CellFormat.Shading.BackgroundPatternColor = Color.Green;
            builder.Writeln("Cell #2");

            // End this row.
            builder.EndRow();

            // Clear the cell formatting from previous operations.
            builder.CellFormat.ClearFormatting();

            // Create the second row.
            builder.InsertCell();

            // Create larger borders for the first cell of this row. This will be different
            // compared to the borders set for the table.
            builder.CellFormat.Borders.Left.LineWidth = 4.0;
            builder.CellFormat.Borders.Right.LineWidth = 4.0;
            builder.CellFormat.Borders.Top.LineWidth = 4.0;
            builder.CellFormat.Borders.Bottom.LineWidth = 4.0;
            builder.Writeln("Cell #3");

            builder.InsertCell();
            // Clear the cell formatting from the previous cell.
            builder.CellFormat.ClearFormatting();
            builder.Writeln("Cell #4");

            doc.Save(MyDir + @"\Artifacts\Table.SetBordersAndShading.doc");
            //ExEnd

            // Verify the table was created correctly.
            Assert.AreEqual(Color.Red.ToArgb(), table.FirstRow.FirstCell.CellFormat.Shading.BackgroundPatternColor.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.Cells[1].CellFormat.Shading.BackgroundPatternColor.ToArgb());
            Assert.AreEqual(Color.Green.ToArgb(), table.FirstRow.Cells[1].CellFormat.Shading.BackgroundPatternColor.ToArgb());
            Assert.AreEqual(Color.Empty.ToArgb(), table.LastRow.FirstCell.CellFormat.Shading.BackgroundPatternColor.ToArgb());

            Assert.AreEqual(Color.Black.ToArgb(), table.FirstRow.FirstCell.CellFormat.Borders.Left.Color.ToArgb()); Assert.AreEqual(Color.Black.ToArgb(), table.FirstRow.FirstCell.CellFormat.Borders.Left.Color.ToArgb());
            Assert.AreEqual(LineStyle.Single, table.FirstRow.FirstCell.CellFormat.Borders.Left.LineStyle);
            Assert.AreEqual(2.0, table.FirstRow.FirstCell.CellFormat.Borders.Left.LineWidth);
            Assert.AreEqual(4.0, table.LastRow.FirstCell.CellFormat.Borders.Left.LineWidth);
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
            //ExSummary:Inserts a hyperlink referencing local bookmark.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartBookmark("Bookmark1");
            builder.Write("Bookmarked text.");
            builder.EndBookmark("Bookmark1");

            builder.Writeln("Some other text");

            // Specify font formatting for the hyperlink.
            builder.Font.Color = Color.Blue;
            builder.Font.Underline = Underline.Single;

            // Insert hyperlink.
            // Switch \o is used to provide hyperlink tip text.
            builder.InsertHyperlink("Hyperlink Text", @"Bookmark1"" \o ""Hyperlink Tip", true);

            // Clear hyperlink formatting.
            builder.Font.ClearFormatting();

            doc.Save(MyDir + @"\Artifacts\DocumentBuilder.InsertHyperlinkToLocalBookmark.doc");
            //ExEnd
        }

        [Test]
        public void DocumentBuilderCtor()
        {
            //ExStart
            //ExId:DocumentBuilderCtor
            //ExSummary:Shows how to create a simple document using a document builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Hello World!");
            //ExEnd
        }

        [Test]
        public void DocumentBuilderCursorPosition()
        {
            //ExStart
            //ExId:DocumentBuilderCursorPosition
            //ExSummary:Shows how to access the current node in a document builder.
            Document doc = new Document(MyDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            Node curNode = builder.CurrentNode;
            Paragraph curParagraph = builder.CurrentParagraph;
            //ExEnd
        }

        [Test]
        public void DocumentBuilderMoveToNode()
        {
            //ExStart
            //ExFor:Story.LastParagraph
            //ExFor:DocumentBuilder.MoveTo(Node)
            //ExId:DocumentBuilderMoveToNode
            //ExSummary:Shows how to move a cursor position to a specified node.
            Document doc = new Document(MyDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveTo(doc.FirstSection.Body.LastParagraph);
            //ExEnd
        }

        [Test]
        public void DocumentBuilderMoveToDocumentStartEnd()
        {
            //ExStart
            //ExId:DocumentBuilderMoveToDocumentStartEnd
            //ExSummary:Shows how to move a cursor position to the beginning or end of a document.
            Document doc = new Document(MyDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToDocumentEnd();
            builder.Writeln("This is the end of the document.");

            builder.MoveToDocumentStart();
            builder.Writeln("This is the beginning of the document.");
            //ExEnd
        }

        [Test]
        public void DocumentBuilderMoveToSection()
        {
            //ExStart
            //ExId:DocumentBuilderMoveToSection
            //ExSummary:Shows how to move a cursor position to the specified section.
            Document doc = new Document(MyDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Parameters are 0-index. Moves to third section.
            builder.MoveToSection(2);
            builder.Writeln("This is the 3rd section.");
            //ExEnd
        }

        [Test]
        public void DocumentBuilderMoveToParagraph()
        {
            //ExStart
            //ExFor:DocumentBuilder.MoveToParagraph
            //ExId:DocumentBuilderMoveToParagraph
            //ExSummary:Shows how to move a cursor position to the specified paragraph.
            Document doc = new Document(MyDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Parameters are 0-index. Moves to third paragraph.
            builder.MoveToParagraph(2, 0);
            builder.Writeln("This is the 3rd paragraph.");
            //ExEnd
        }

        [Test]
        public void DocumentBuilderMoveToTableCell()
        {
            //ExStart
            //ExFor:DocumentBuilder.MoveToCell
            //ExId:DocumentBuilderMoveToTableCell
            //ExSummary:Shows how to move a cursor position to the specified table cell.
            Document doc = new Document(MyDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // All parameters are 0-index. Moves to the 2nd table, 3rd row, 5th cell.
            builder.MoveToCell(1, 2, 4, 0);
            builder.Writeln("Hello World!");
            //ExEnd
        }

        [Test]
        public void DocumentBuilderMoveToBookmark()
        {
            //ExStart
            //ExId:DocumentBuilderMoveToBookmark
            //ExSummary:Shows how to move a cursor position to a bookmark.
            Document doc = new Document(MyDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToBookmark("CoolBookmark");
            builder.Writeln("This is a very cool bookmark.");
            //ExEnd
        }

        [Test]
        public void DocumentBuilderMoveToBookmarkEnd()
        {
            //ExStart
            //ExFor:DocumentBuilder.MoveToBookmark(String, Boolean, Boolean)
            //ExId:DocumentBuilderMoveToBookmarkEnd
            //ExSummary:Shows how to move a cursor position to just after the bookmark end.
            Document doc = new Document(MyDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToBookmark("CoolBookmark", false, true);
            builder.Writeln("This is a very cool bookmark.");
            //ExEnd
        }

        [Test]
        public void DocumentBuilderMoveToMergeField()
        {
            //ExStart
            //ExId:DocumentBuilderMoveToMergeField
            //ExSummary:Shows how to move the cursor to a position just beyond the specified merge field.
            Document doc = new Document(MyDir + "DocumentBuilder.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToMergeField("NiceMergeField");
            builder.Writeln("This is a very nice merge field.");
            //ExEnd
        }

        [Test]
        public void DocumentBuilderInsertParagraph()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertParagraph
            //ExFor:ParagraphFormat.FirstLineIndent
            //ExFor:ParagraphFormat.Alignment
            //ExFor:ParagraphFormat.KeepTogether
            //ExId:DocumentBuilderInsertParagraph
            //ExSummary:Shows how to insert a paragraph into the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Specify font formatting
            Aspose.Words.Font font = builder.Font;
            font.Size = 16;
            font.Bold = true;
            font.Color = Color.Blue;
            font.Name = "Arial";
            font.Underline = Underline.Dash;

            // Specify paragraph formatting
            ParagraphFormat paragraphFormat = builder.ParagraphFormat;
            paragraphFormat.FirstLineIndent = 8;
            paragraphFormat.Alignment = ParagraphAlignment.Justify;
            paragraphFormat.KeepTogether = true;

            builder.Writeln("A whole paragraph.");
            //ExEnd
        }

        [Test]
        public void DocumentBuilderBuildTable()
        {
            //ExStart
            //ExFor:Table
            //ExFor:DocumentBuilder.StartTable
            //ExFor:DocumentBuilder.InsertCell
            //ExFor:DocumentBuilder.EndRow
            //ExFor:DocumentBuilder.EndTable
            //ExFor:DocumentBuilder.CellFormat
            //ExFor:DocumentBuilder.RowFormat
            //ExFor:DocumentBuilder.Write
            //ExFor:DocumentBuilder.Writeln(String)
            //ExFor:RowFormat.Height
            //ExFor:RowFormat.HeightRule
            //ExFor:CellVerticalAlignment
            //ExFor:CellFormat.Orientation
            //ExFor:TextOrientation
            //ExFor:Table.AutoFit
            //ExFor:AutoFitBehavior
            //ExId:DocumentBuilderBuildTable
            //ExSummary:Shows how to build a formatted table that contains 2 rows and 2 columns.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();

            // Insert a cell
            builder.InsertCell();
            // Use fixed column widths.
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            builder.CellFormat.VerticalAlignment = CellVerticalAlignment.Center;
            builder.Write("This is row 1 cell 1");

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
            builder.Writeln("This is row 2 cell 1");

            // Insert a cell
            builder.InsertCell();
            builder.CellFormat.Orientation = TextOrientation.Downward;
            builder.Writeln("This is row 2 cell 2");

            builder.EndRow();

            builder.EndTable();
            //ExEnd
        }

        [Test]
        public void TableCellVerticalRotatedFarEastTextOrientation()
        {
            Document doc = new Document(MyDir + "DocumentBuilder.TableCellVerticalRotatedFarEastTextOrientation.docx");

            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            Cell cell = table.FirstRow.FirstCell;

            Assert.AreEqual(TextOrientation.VerticalRotatedFarEast, cell.CellFormat.Orientation);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            table = (Table)doc.GetChild(NodeType.Table, 0, true);
            cell = table.FirstRow.FirstCell;

            Assert.AreEqual(TextOrientation.VerticalRotatedFarEast, cell.CellFormat.Orientation);
        }

        [Test]
        public void DocumentBuilderInsertBreak()
        {
            //ExStart
            //ExId:DocumentBuilderInsertBreak
            //ExSummary:Shows how to insert page breaks into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("This is page 1.");
            builder.InsertBreak(BreakType.PageBreak);

            builder.Writeln("This is page 2.");
            builder.InsertBreak(BreakType.PageBreak);

            builder.Writeln("This is page 3.");
            //ExEnd
        }

        [Test]
        public void DocumentBuilderInsertInlineImage()
        {
            //ExStart
            //ExId:DocumentBuilderInsertInlineImage
            //ExSummary:Shows how to insert an inline image at the cursor position into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage(MyDir + "Watermark.png");
            //ExEnd
        }

        [Test]
        public void DocumentBuilderInsertFloatingImage()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExId:DocumentBuilderInsertFloatingImage
            //ExSummary:Shows how to insert a floating image from a file or URL.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage(MyDir + "Watermark.png",
                RelativeHorizontalPosition.Margin,
                100,
                RelativeVerticalPosition.Margin,
                100,
                200,
                100,
                WrapType.Square);
            //ExEnd
        }

        [Test]
        public void InsertImageFromUrl()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(String)
            //ExSummary:Shows how to insert an image into a document from a web address.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage("http://www.aspose.com/images/aspose-logo.gif");

            doc.Save(MyDir + @"\Artifacts\DocumentBuilder.InsertImageFromUrl.doc");
            //ExEnd

            // Verify that the image was inserted into the document.
            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            Assert.IsNotNull(shape);
            Assert.True(shape.HasImage);
        }

        [Test]
        public void DocumentBuilderInsertImageSourceSize()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertImage(String, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExId:DocumentBuilderInsertFloatingImageSourceSize
            //ExSummary:Shows how to insert a floating image from a file or URL and retain the original image size in the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Pass a negative value to the width and height values to specify using the size of the source image.
            builder.InsertImage(MyDir + "LogoSmall.png",
                RelativeHorizontalPosition.Margin,
                200,
                RelativeVerticalPosition.Margin,
                100,
                -1,
                -1,
                WrapType.Square);
            //ExEnd

            doc.Save(MyDir + @"\Artifacts\DocumentBuilder.InsertImageOriginalSize.doc");
        }

        [Test]
        public void DocumentBuilderInsertBookmark()
        {
            //ExStart
            //ExId:DocumentBuilderInsertBookmark
            //ExSummary:Shows how to insert a bookmark into a document using a document builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartBookmark("FineBookmark");
            builder.Writeln("This is just a fine bookmark.");
            builder.EndBookmark("FineBookmark");
            //ExEnd
        }

        [Test]
        public void DocumentBuilderInsertTextInputFormField()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertTextInput
            //ExId:DocumentBuilderInsertTextInputFormField
            //ExSummary:Shows how to insert a text input form field into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Hello", 0);
            //ExEnd
        }

        [Test]
        public void DocumentBuilderInsertCheckBoxFormField()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertCheckBox
            //ExId:DocumentBuilderInsertCheckBoxFormField
            //ExSummary:Shows how to insert a checkbox form field into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertCheckBox("CheckBox", true, 0);
            //ExEnd
        }

        [Test]
        public void DocumentBuilderInsertComboBoxFormField()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertComboBox
            //ExId:DocumentBuilderInsertComboBoxFormField
            //ExSummary:Shows how to insert a combobox form field into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            string[] items = { "One", "Two", "Three" };
            builder.InsertComboBox("DropDown", items, 0);
            //ExEnd
        }

        [Test]
        public void DocumentBuilderInsertToc()
        {
            //ExStart
            //ExId:DocumentBuilderInsertTOC
            //ExSummary:Shows how to insert a Table of Contents field into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table of contents at the beginning of the document.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

            // The newly inserted table of contents will be initially empty.
            // It needs to be populated by updating the fields in the document.
            doc.UpdateFields();
            //ExEnd
        }

        [Test]
        public void InsertSignatureLine()
        {
            
        }
        
        [Test]
        public void InsertSignatureLineCurrentPozition()
        {
            //ExStart
            //ExFor:SignatureLine
            //ExFor:SignatureLineOptions
            //ExFor:DocumentBuilder.InsertSignatureLine(SignatureLineOptions)
            //ExFor:DocumentBuilder.InsertSignatureLine(SignatureLineOptions)
            //ExSummary:Shows how to insert signature line and get signature line properties
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            SignatureLineOptions options = new SignatureLineOptions();
            options.Signer = "John Doe";
            options.SignerTitle = "Manager";
            options.Email = "johndoe@aspose.com";
            options.ShowDate = true;
            options.DefaultInstructions = false;
            options.Instructions = "You need more info about signature line";
            options.AllowComments = true;

            builder.InsertSignatureLine(options);
            builder.InsertSignatureLine(options, RelativeHorizontalPosition.RightMargin, 2.0, RelativeVerticalPosition.Page, 3.0, WrapType.Inline);
            
            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
            SignatureLine signatureLine = shape.SignatureLine;

            Assert.AreEqual("John Doe", signatureLine.Signer);
            Assert.AreEqual("Manager", signatureLine.SignerTitle);
            Assert.AreEqual("johndoe@aspose.com", signatureLine.Email);
            Assert.AreEqual(true, signatureLine.ShowDate);
            Assert.AreEqual(false, signatureLine.DefaultInstructions);
            Assert.AreEqual("You need more info about signature line", signatureLine.Instructions);
            Assert.AreEqual(true, signatureLine.AllowComments);
            Assert.AreEqual(false, signatureLine.IsSigned);
            Assert.AreEqual(false, signatureLine.IsValid);
            //ExEnd

            shape = (Shape)doc.GetChild(NodeType.Shape, 1, true);
            Assert.AreEqual(RelativeHorizontalPosition.RightMargin, shape.RelativeHorizontalPosition);
            Assert.AreEqual(2.0, shape.Left);
            Assert.AreEqual(RelativeVerticalPosition.Page, shape.RelativeVerticalPosition);
            Assert.AreEqual(3.0, shape.Top);
            Assert.AreEqual(WrapType.Inline, shape.WrapType);
            //Bug: If wraptype are not inline shape break his position (builder.InsertSignatureLine(options, RelativeHorizontalPosition.RightMargin, 2.0, RelativeVerticalPosition.Page, 3.0, WrapType.Inline);)
        }

        [Test]
        public void DocumentBuilderSetFontFormatting()
        {
            //ExStart
            //ExId:DocumentBuilderSetFontFormatting
            //ExSummary:Shows how to set font formatting.
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
            builder.Writeln("I'm a very nice formatted string.");
            //ExEnd
        }

        [Test]
        public void DocumentBuilderSetParagraphFormatting()
        {
            //ExStart
            //ExFor:ParagraphFormat.RightIndent
            //ExFor:ParagraphFormat.LeftIndent
            //ExFor:ParagraphFormat.SpaceAfter
            //ExId:DocumentBuilderSetParagraphFormatting
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
            builder.Writeln("I'm a very nice formatted paragraph. I'm intended to demonstrate how the left and right indents affect word wrapping.");
            builder.Writeln("I'm another nice formatted paragraph. I'm intended to demonstrate how the space after paragraph looks like.");
            //ExEnd
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
            //ExId:DocumentBuilderSetCellFormatting
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

            builder.Writeln("I'm a wonderful formatted cell.");

            builder.EndRow();
            builder.EndTable();
            //ExEnd
        }

        [Test]
        public void DocumentBuilderSetRowFormatting()
        {
            //ExStart
            //ExFor:DocumentBuilder.RowFormat
            //ExFor:RowFormat.Height
            //ExFor:RowFormat.HeightRule
            //ExFor:Table.LeftPadding
            //ExFor:Table.RightPadding
            //ExFor:Table.TopPadding
            //ExFor:Table.BottomPadding
            //ExId:DocumentBuilderSetRowFormatting
            //ExSummary:Shows how to create a table that contains a single cell and apply row formatting.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.InsertCell();

            // Set the row formatting
            RowFormat rowFormat = builder.RowFormat;
            rowFormat.Height = 100;
            rowFormat.HeightRule = HeightRule.Exactly;
            // These formatting properties are set on the table and are applied to all rows in the table.
            table.LeftPadding = 30;
            table.RightPadding = 30;
            table.TopPadding = 30;
            table.BottomPadding = 30;

            builder.Writeln("I'm a wonderful formatted row.");

            builder.EndRow();
            builder.EndTable();
            //ExEnd
        }

        [Test]
        public void DocumentBuilderSetListFormatting()
        {
            //ExStart
            //ExId:DocumentBuilderSetListFormatting
            //ExSummary:Shows how to build a multilevel list.
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
            //ExEnd
        }

        [Test]
        public void DocumentBuilderSetSectionFormatting()
        {
            //ExStart
            //ExId:DocumentBuilderSetSectionFormatting
            //ExSummary:Shows how to set such properties as page size and orientation for the current section.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set page properties
            builder.PageSetup.Orientation = Orientation.Landscape;
            builder.PageSetup.LeftMargin = 50;
            builder.PageSetup.PaperSize = PaperSize.Paper10x14;
            //ExEnd
        }

        [Test]
        public void InsertFootnote()
        {
            //ExStart
            //ExFor:Footnote
            //ExFor:FootnoteType
            //ExFor:DocumentBuilder.InsertFootnote(FootnoteType,string)
            //ExFor:DocumentBuilder.InsertFootnote(FootnoteType,string,string)
            //ExSummary:Shows how to add a footnote to a paragraph in the document using DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Some text");

            builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.");
            builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.", "242");
            //ExEnd

            Assert.AreEqual("Footnote text.", doc.GetChildNodes(NodeType.Footnote, true)[0].ToString(SaveFormat.Text).Trim());
        }

        [Test]
        public void AddFootnoteWithCustomMarks()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Some text");

            Footnote foot = new Footnote(doc, FootnoteType.Footnote);
            foot.ReferenceMark = "242";

            builder.InsertFootnote(FootnoteType.Footnote, "Footnote text.", foot.ReferenceMark);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            doc = new Document(dstStream);
            foot = (Footnote)doc.GetChildNodes(NodeType.Footnote, true)[0];
            
            Assert.IsFalse(foot.IsAuto);
            Assert.AreEqual("242", foot.ReferenceMark);
            Assert.AreEqual("242 Footnote text.\r", foot.GetText());
        }

        [Test]
        public void DocumentBuilderApplyParagraphStyle()
        {
            //ExStart
            //ExId:DocumentBuilderApplyParagraphStyle
            //ExSummary:Shows how to apply a paragraph style.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set paragraph style
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;

            builder.Write("Hello");
            //ExEnd
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
            //ExId:DocumentBuilderApplyBordersAndShading
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

            builder.Write("I'm a formatted paragraph with double border and nice shading.");
            //ExEnd
        }

        [Test]
        public void DeleteRowEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.DeleteRow
            //ExSummary:Shows how to delete a row from a table.
            Document doc = new Document(MyDir + "DocumentBuilder.DocWithTable.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Delete the first row of the first table in the document.
            builder.DeleteRow(0, 0);
            //ExEnd
        }

        //ToDo: There is some unclear behavior
        [Test]
        public void InsertDocumentEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertDocument
            //ExSummary:Shows how to insert a document into another document.
            Document doc = new Document(MyDir + "Document.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);
            Document docToInsert = new Document(MyDir + "DocumentBuilder.InsertedDoc.doc");

            builder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);
            //ExEnd
        }

        [Test]
        public void MoveToFieldEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.MoveToField
            //ExSummary:Shows how to move document builder's cursor to a specific field.
            Document doc = new Document(MyDir + "Document.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            Field field = builder.InsertField("MERGEFIELD field");

            builder.MoveToField(field, true);
            //ExEnd
        }

        [Test]
        public void InsertOleObjectEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertOleObject(String, Boolean, Boolean, Image)
            //ExFor:DocumentBuilder.InsertOleObject(String, String, Boolean, Boolean, Image)
            //ExSummary:Shows how to insert an OLE object into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            Image representingImage = Image.FromFile(MyDir + "Aspose.Words.gif");
            
            Shape oleObject = builder.InsertOleObject(MyDir + "Document.Spreadsheet.xlsx", false, false, representingImage);
            Shape oleObjectProgId = builder.InsertOleObject("http://www.aspose.com", "htmlfile", true, true, null);

            // Double click on the image in the .doc to see the spreadsheet.
            // Double click on the icon in the .doc to see the html.
            doc.Save(MyDir + @"\Artifacts\Document.InsertedOleObject.doc");
            //ExEnd

            //ToDo: There is some bug, need more info for this (breaking html link)
            //Shape oleObjectProgId = builder.InsertOleObject("http://www.aspose.com", "htmlfile", true, false, null);
        }

        [Test]
        public void InsertOleObjectException()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Assert.That(() => builder.InsertOleObject("", "checkbox", false, true, null), Throws.TypeOf<ArgumentException>());
        }

        [Test]
        public void InsertChartDoubleEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertChart(ChartType, Double, Double)
            //ExSummary:Shows how to insert a chart into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertChart(ChartType.Pie, ConvertUtil.PixelToPoint(300),
                                ConvertUtil.PixelToPoint(300));

            doc.Save(MyDir + @"\Artifacts\Document.InsertedChartDouble.doc");
            //ExEnd
        }

        [Test]
        public void DataArraysWrongSize()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add chart with default data.
            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);
            Chart chart = shape.Chart;

            ChartSeriesCollection seriesColl = chart.Series;
            seriesColl.Clear();

            // Create category names array, second category will be null.
            string[] categories = new string[] { "Cat1", null, "Cat3", "Cat4", "Cat5", null };

            // Adding new series with empty (double.NaN) values.
            seriesColl.Add("AW Series 1", categories, new double[] { 1, 2, double.NaN, 4, 5, 6 });
            seriesColl.Add("AW Series 2", categories, new double[] { 2, 3, double.NaN, 5, 6, 7 });
            Assert.That(() => seriesColl.Add("AW Series 3", categories, new double[] { double.NaN, 4, 5, double.NaN, double.NaN }), Throws.TypeOf<ArgumentException>());
            Assert.That(() => seriesColl.Add("AW Series 4", categories, new double[] { double.NaN, double.NaN, double.NaN, double.NaN, double.NaN }), Throws.TypeOf<ArgumentException>());
        }
        
        [Test]
        public void EmptyValuesInChartData()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            // Add chart with default data.
            Shape shape = builder.InsertChart(ChartType.Line, 432, 252);
            Chart chart = shape.Chart;

            ChartSeriesCollection seriesColl = chart.Series;
            seriesColl.Clear();

            // Create category names array, second category will be null.
            string[] categories = new string[] { "Cat1", null, "Cat3", "Cat4", "Cat5", null };

            // Adding new series with empty (double.NaN) values.
            seriesColl.Add("AW Series 1", categories, new double[] { 1, 2, double.NaN, 4, 5, 6 });
            seriesColl.Add("AW Series 2", categories, new double[] { 2, 3, double.NaN, 5, 6, 7 });
            seriesColl.Add("AW Series 3", categories, new double[] { double.NaN, 4, 5, double.NaN, 7, 8 });
            seriesColl.Add("AW Series 4", categories, new double[] { double.NaN, double.NaN, double.NaN, double.NaN, double.NaN, 9 });

            doc.Save(MyDir + @"\Artifacts\EmptyValuesInChartData.docx");
        }

        [Test]
        public void InsertChartRelativePositionEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertChart(ChartType, RelativeHorizontalPosition, Double, RelativeVerticalPosition, Double, Double, Double, WrapType)
            //ExSummary:Shows how to insert a chart into a document and specify all positioning options in the arguments.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertChart(ChartType.Pie, RelativeHorizontalPosition.Margin, 100, RelativeVerticalPosition.Margin, 100,
                                    200, 100, WrapType.Square);

            doc.Save(MyDir + @"\Artifacts\Document.InsertedChartRelativePosition.doc");
            //ExEnd
        }

        [Test]
        public void InsertCheckBoxEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertCheckBox(String, Boolean, Int32)
            //ExFor:DocumentBuilder.InsertCheckBox(String, Boolean, Boolean, Int32)
            //ExSummary:Shows how to insert a check box into a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a checkbox with no default value and let MS Word apply the default size.
            builder.Writeln("Check box 1");
            builder.InsertCheckBox("CheckBox1", false, 0);
            builder.Writeln();

            // Insert a checked checkbox with a specified value.
            builder.Writeln("Check box 2");
            builder.InsertCheckBox("CheckBox2", false, true, 50);

            doc.Save(MyDir + @"\Artifacts\Document.InsertedCheckBoxes.doc");
            //ExEnd
        }

        [Test]
        public void InsertFieldEx()
        {
            //ExStart
            //ExFor:DocumentBuilder.InsertField(FieldType, Boolean)
            //ExSummary:Shows how to insert a field.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("This field was inserted/updated at ");
            builder.InsertField(FieldType.FieldTime, true);

            doc.Save(MyDir + @"\Artifacts\Document.InsertedField.doc");
            //ExEnd
        }
    }
}