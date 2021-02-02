// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;
using ColorMode = Aspose.Words.Saving.ColorMode;
using Document = Aspose.Words.Document;
using IWarningCallback = Aspose.Words.IWarningCallback;
using PdfSaveOptions = Aspose.Words.Saving.PdfSaveOptions;
using SaveFormat = Aspose.Words.SaveFormat;
using SaveOptions = Aspose.Words.Saving.SaveOptions;
using WarningInfo = Aspose.Words.WarningInfo;
using WarningType = Aspose.Words.WarningType;
using Image =
#if NET462 || JAVA
System.Drawing.Image;
#elif NETCOREAPP2_1 || __MOBILE__
SkiaSharp.SKBitmap;
using SkiaSharp;
#endif
#if NET462 || NETCOREAPP2_1 || JAVA
using Aspose.Pdf;
using Aspose.Pdf.Annotations;
using Aspose.Pdf.Facades;
using Aspose.Pdf.Forms;
using Aspose.Pdf.Operators;
using Aspose.Pdf.Text;
#endif

namespace ApiExamples
{
    [TestFixture]
    internal class ExPdfSaveOptions : ApiExampleBase
    {
        [Test]
        public void OnePage()
        {
            //ExStart
            //ExFor:FixedPageSaveOptions.PageSet
            //ExFor:Document.Save(Stream, SaveOptions)
            //ExSummary:Shows how to convert only some of the pages in a document to PDF.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 3.");

            using (Stream stream = File.Create(ArtifactsDir + "PdfSaveOptions.OnePage.pdf"))
            {
                // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
                // to modify how that method converts the document to .PDF.
                PdfSaveOptions options = new PdfSaveOptions();

                // Set the "PageIndex" to "1" to render a portion of the document starting from the second page.
                options.PageSet = new PageSet(1);

                // This document will contain one page starting from page two, which will only contain the second page.
                doc.Save(stream, options);
            }
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.OnePage.pdf");

            Assert.AreEqual(1, pdfDocument.Pages.Count);

            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
            pdfDocument.Pages.Accept(textFragmentAbsorber);

            Assert.AreEqual("Page 2.", textFragmentAbsorber.Text);
#endif
        }

        [Test]
        public void HeadingsOutlineLevels()
        {
            //ExStart
            //ExFor:ParagraphFormat.IsHeading
            //ExFor:PdfSaveOptions.OutlineOptions
            //ExFor:PdfSaveOptions.SaveFormat
            //ExSummary:Shows how to limit the headings' level that will appear in the outline of a saved PDF document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert headings that can serve as TOC entries of levels 1, 2, and then 3.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            Assert.True(builder.ParagraphFormat.IsHeading);

            builder.Writeln("Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

            builder.Writeln("Heading 1.1");
            builder.Writeln("Heading 1.2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;

            builder.Writeln("Heading 1.2.1");
            builder.Writeln("Heading 1.2.2");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.SaveFormat = SaveFormat.Pdf;

            // The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
            // Clicking on an entry in this outline will take us to the location of its respective heading.
            // Set the "HeadingsOutlineLevels" property to "2" to exclude all headings whose levels are above 2 from the outline.
            // The last two headings we have inserted above will not appear.
            saveOptions.OutlineOptions.HeadingsOutlineLevels = 2;

            doc.Save(ArtifactsDir + "PdfSaveOptions.HeadingsOutlineLevels.pdf", saveOptions);
            //ExEnd

            #if NET462 || NETCOREAPP2_1 || JAVA
            PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
            bookmarkEditor.BindPdf(ArtifactsDir + "PdfSaveOptions.HeadingsOutlineLevels.pdf");

            Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

            Assert.AreEqual(3, bookmarks.Count);
            #endif
        }

        [TestCase(false)]
        [TestCase(true)]
        public void CreateMissingOutlineLevels(bool createMissingOutlineLevels)
        {
            //ExStart
            //ExFor:OutlineOptions.CreateMissingOutlineLevels
            //ExFor:PdfSaveOptions.OutlineOptions
            //ExSummary:Shows how to work with outline levels that do not contain any corresponding headings when saving a PDF document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert headings that can serve as TOC entries of levels 1 and 5.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            Assert.True(builder.ParagraphFormat.IsHeading);

            builder.Writeln("Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading5;

            builder.Writeln("Heading 1.1.1.1.1");
            builder.Writeln("Heading 1.1.1.1.2");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
            // Clicking on an entry in this outline will take us to the location of its respective heading.
            // Set the "HeadingsOutlineLevels" property to "5" to include all headings of levels 5 and below in the outline.
            saveOptions.OutlineOptions.HeadingsOutlineLevels = 5;

            // This document contains headings of levels 1 and 5, and no headings with levels of 2, 3, and 4. 
            // The output PDF document will treat outline levels 2, 3, and 4 as "missing".
            // Set the "CreateMissingOutlineLevels" property to "true" to include all missing levels in the outline,
            // leaving blank outline entries since there are no usable headings.
            // Set the "CreateMissingOutlineLevels" property to "false" to ignore missing outline levels,
            // and treat the outline level 5 headings as level 2.
            saveOptions.OutlineOptions.CreateMissingOutlineLevels = createMissingOutlineLevels;

            doc.Save(ArtifactsDir + "PdfSaveOptions.CreateMissingOutlineLevels.pdf", saveOptions);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            PdfBookmarkEditor bookmarkEditor = new PdfBookmarkEditor();
            bookmarkEditor.BindPdf(ArtifactsDir + "PdfSaveOptions.CreateMissingOutlineLevels.pdf");

            Bookmarks bookmarks = bookmarkEditor.ExtractBookmarks();

            Assert.AreEqual(createMissingOutlineLevels ? 6 : 3, bookmarks.Count);
#endif
        }

        [TestCase(false)]
        [TestCase(true)]
        public void TableHeadingOutlines(bool createOutlinesForHeadingsInTables)
        {
            //ExStart
            //ExFor:OutlineOptions.CreateOutlinesForHeadingsInTables
            //ExSummary:Shows how to create PDF document outline entries for headings inside tables.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a table with three rows. The first row,
            // whose text we will format in a heading-type style, will serve as the column header.
            builder.StartTable();
            builder.InsertCell();
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Write("Customers");
            builder.EndRow();
            builder.InsertCell();
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Write("John Doe");
            builder.EndRow();
            builder.InsertCell();
            builder.Write("Jane Doe");
            builder.EndTable();

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
            // Clicking on an entry in this outline will take us to the location of its respective heading.
            // Set the "HeadingsOutlineLevels" property to "1" to get the outline
            // to only register headings with heading levels that are no larger than 1.
            pdfSaveOptions.OutlineOptions.HeadingsOutlineLevels = 1;

            // Set the "CreateOutlinesForHeadingsInTables" property to "false" to exclude all headings within tables,
            // such as the one we have created above from the outline.
            // Set the "CreateOutlinesForHeadingsInTables" property to "true" to include all headings within tables
            // in the outline, provided that they have a heading level that is no larger than the value of the "HeadingsOutlineLevels" property.
            pdfSaveOptions.OutlineOptions.CreateOutlinesForHeadingsInTables = createOutlinesForHeadingsInTables;

            doc.Save(ArtifactsDir + "PdfSaveOptions.TableHeadingOutlines.pdf", pdfSaveOptions);
            //ExEnd

            #if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.TableHeadingOutlines.pdf");

            if (createOutlinesForHeadingsInTables)
            {
                Assert.AreEqual(1, pdfDoc.Outlines.Count);
                Assert.AreEqual("Customers", pdfDoc.Outlines[1].Title);
            } else
                Assert.AreEqual(0, pdfDoc.Outlines.Count);

            TableAbsorber tableAbsorber = new TableAbsorber();
            tableAbsorber.Visit(pdfDoc.Pages[1]);

            Assert.AreEqual("Customers", tableAbsorber.TableList[0].RowList[0].CellList[0].TextFragments[1].Text);
            Assert.AreEqual("John Doe", tableAbsorber.TableList[0].RowList[1].CellList[0].TextFragments[1].Text);
            Assert.AreEqual("Jane Doe", tableAbsorber.TableList[0].RowList[2].CellList[0].TextFragments[1].Text);
#endif
        }

        [Test]
        public void ExpandedOutlineLevels()
        {
            //ExStart
            //ExFor:Document.Save(String, SaveOptions)
            //ExFor:PdfSaveOptions
            //ExFor:OutlineOptions.HeadingsOutlineLevels
            //ExFor:OutlineOptions.ExpandedOutlineLevels
            //ExSummary:Shows how to convert a whole document to PDF with three levels in the document outline.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert headings of levels 1 to 5.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            Assert.True(builder.ParagraphFormat.IsHeading);

            builder.Writeln("Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

            builder.Writeln("Heading 1.1");
            builder.Writeln("Heading 1.2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;

            builder.Writeln("Heading 1.2.1");
            builder.Writeln("Heading 1.2.2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading4;

            builder.Writeln("Heading 1.2.2.1");
            builder.Writeln("Heading 1.2.2.2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading5;

            builder.Writeln("Heading 1.2.2.2.1");
            builder.Writeln("Heading 1.2.2.2.2");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // The output PDF document will contain an outline, which is a table of contents that lists headings in the document body.
            // Clicking on an entry in this outline will take us to the location of its respective heading.
            // Set the "HeadingsOutlineLevels" property to "4" to exclude all headings whose levels are above 4 from the outline.
            options.OutlineOptions.HeadingsOutlineLevels = 4;

            // If an outline entry has subsequent entries of a higher level inbetween itself and the next entry of the same or lower level,
            // an arrow will appear to the left of the entry. This entry is the "owner" of several such "sub-entries".
            // In our document, the outline entries from the 5th heading level are sub-entries of the second 4th level outline entry,
            // the 4th and 5th heading level entries are sub-entries of the second 3rd level entry, and so on. 
            // In the outline, we can click on the arrow of the "owner" entry to collapse/expand all its sub-entries.
            // Set the "ExpandedOutlineLevels" property to "2" to automatically expand all heading level 2 and lower outline entries
            // and collapse all level and 3 and higher entries when we open the document. 
            options.OutlineOptions.ExpandedOutlineLevels = 2;

            doc.Save(ArtifactsDir + "PdfSaveOptions.ExpandedOutlineLevels.pdf", options);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.ExpandedOutlineLevels.pdf");

            Assert.AreEqual(1, pdfDocument.Outlines.Count);
            Assert.AreEqual(5, pdfDocument.Outlines.VisibleCount);

            Assert.True(pdfDocument.Outlines[1].Open);
            Assert.AreEqual(1, pdfDocument.Outlines[1].Level);

            Assert.False(pdfDocument.Outlines[1][1].Open);
            Assert.AreEqual(2, pdfDocument.Outlines[1][1].Level);

            Assert.True(pdfDocument.Outlines[1][2].Open);
            Assert.AreEqual(2, pdfDocument.Outlines[1][2].Level);
#endif
        }

        [TestCase(false)]
        [TestCase(true)]
        public void UpdateFields(bool updateFields)
        {
            //ExStart
            //ExFor:PdfSaveOptions.Clone
            //ExFor:SaveOptions.UpdateFields
            //ExSummary:Shows how to update all the fields in a document immediately before saving it to PDF.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert text with PAGE and NUMPAGES fields. These fields do not display the correct value in real time.
            // We will need to manually update them using updating methods such as "Field.Update()", and "Document.UpdateFields()"
            // each time we need them to display accurate values.
            builder.Write("Page ");
            builder.InsertField("PAGE", "");
            builder.Write(" of ");
            builder.InsertField("NUMPAGES", "");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Hello World!");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Set the "UpdateFields" property to "false" to not update all the fields in a document right before a save operation.
            // This is the preferable option if we know that all our fields will be up to date before saving.
            // Set the "UpdateFields" property to "true" to iterate through all the document
            // fields and update them before we save it as a PDF. This will make sure that all the fields will display
            // the most accurate values in the PDF.
            options.UpdateFields = updateFields;
            
            // We can clone PdfSaveOptions objects.
            Assert.AreNotSame(options, options.Clone());

            doc.Save(ArtifactsDir + "PdfSaveOptions.UpdateFields.pdf", options);
            //ExEnd

            #if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.UpdateFields.pdf");

            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
            pdfDocument.Pages.Accept(textFragmentAbsorber);

            Assert.AreEqual(updateFields ? "Page 1 of 2" : "Page  of ", textFragmentAbsorber.TextFragments[1].Text);
#endif
        }

        [TestCase(false)]
        [TestCase(true)]
        public void PreserveFormFields(bool preserveFormFields)
        {
            //ExStart
            //ExFor:PdfSaveOptions.PreserveFormFields
            //ExSummary:Shows how to save a document to the PDF format using the Save method and the PdfSaveOptions class.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Please select a fruit: ");

            // Insert a combo box which will allow a user to choose an option from a collection of strings.
            builder.InsertComboBox("MyComboBox", new[] { "Apple", "Banana", "Cherry" }, 0);

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set the "PreserveFormFields" property to "true" to save form fields as interactive objects in the output PDF.
            // Set the "PreserveFormFields" property to "false" to freeze all form fields in the document at
            // their current values and display them as plain text in the output PDF.
            pdfOptions.PreserveFormFields = preserveFormFields;

            doc.Save(ArtifactsDir + "PdfSaveOptions.PreserveFormFields.pdf", pdfOptions);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.PreserveFormFields.pdf");

            Assert.AreEqual(1, pdfDocument.Pages.Count);

            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
            pdfDocument.Pages.Accept(textFragmentAbsorber);

            if (preserveFormFields)
            {
                Assert.AreEqual("Please select a fruit: ", textFragmentAbsorber.Text);
                TestUtil.FileContainsString("10 0 obj\r\n" +
                                            "<</Type /Annot/Subtype /Widget/P 4 0 R/FT /Ch/F 4/Rect [168.39199829 707.35101318 217.87442017 722.64007568]/Ff 131072/T(þÿ\0M\0y\0C\0o\0m\0b\0o\0B\0o\0x)/Opt " +
                                            "[(þÿ\0A\0p\0p\0l\0e) (þÿ\0B\0a\0n\0a\0n\0a) (þÿ\0C\0h\0e\0r\0r\0y) ]/V(þÿ\0A\0p\0p\0l\0e)/DA(0 g /FAAABC 12 Tf )/AP<</N 11 0 R>>>>",
                    ArtifactsDir + "PdfSaveOptions.PreserveFormFields.pdf");

                Aspose.Pdf.Forms.Form form = pdfDocument.Form;
                Assert.AreEqual(1, pdfDocument.Form.Count);

                ComboBoxField field = (ComboBoxField)form.Fields[0];
                
                Assert.AreEqual("MyComboBox", field.FullName);
                Assert.AreEqual(3, field.Options.Count);
                Assert.AreEqual("Apple", field.Value);
            }
            else
            {
                Assert.AreEqual("Please select a fruit: Apple", textFragmentAbsorber.Text);
                Assert.Throws<AssertionException>(() =>
                {
                    TestUtil.FileContainsString("/Widget",
                        ArtifactsDir + "PdfSaveOptions.PreserveFormFields.pdf");
                });

                Assert.AreEqual(0, pdfDocument.Form.Count);
            }
#endif
        }

        [TestCase(PdfCompliance.PdfA1b)]
        [TestCase(PdfCompliance.Pdf17)]
        [TestCase(PdfCompliance.PdfA1a)]
        public void Compliance(PdfCompliance pdfCompliance)
        {
            //ExStart
            //ExFor:PdfSaveOptions.Compliance
            //ExFor:PdfCompliance
            //ExSummary:Shows how to set the PDF standards compliance level of saved PDF documents.
            Document doc = new Document(MyDir + "Images.docx");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Set the "Compliance" property to "PdfCompliance.PdfA1b" to comply with the "PDF/A-1b" standard,
            // which aims to preserve the visual appearance of the document as Aspose.Words convert it to PDF.
            // Set the "Compliance" property to "PdfCompliance.Pdf17" to comply with the "1.7" standard.
            // Set the "Compliance" property to "PdfCompliance.PdfA1a" to comply with the "PDF/A-1a" standard,
            // which complies with "PDF/A-1b" as well as preserving the document structure of the original document.
            // This helps with making documents searchable but may significantly increase the size of already large documents.
            saveOptions.Compliance = pdfCompliance;

            doc.Save(ArtifactsDir + "PdfSaveOptions.Compliance.pdf", saveOptions);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.Compliance.pdf");

            switch (pdfCompliance)
            {
                case PdfCompliance.Pdf17:
                    Assert.AreEqual(PdfFormat.v_1_7, pdfDocument.PdfFormat);
                    Assert.AreEqual("1.7", pdfDocument.Version);
                    break;
                case PdfCompliance.PdfA1a:
                    Assert.AreEqual(PdfFormat.PDF_A_1A, pdfDocument.PdfFormat);
                    Assert.AreEqual("1.4", pdfDocument.Version);
                    break;
                case PdfCompliance.PdfA1b:
                    Assert.AreEqual(PdfFormat.PDF_A_1B, pdfDocument.PdfFormat);
                    Assert.AreEqual("1.4", pdfDocument.Version);
                    break;
            }
#endif
        }

        [TestCase(PdfTextCompression.None)]
        [TestCase(PdfTextCompression.Flate)]
        public void TextCompression(PdfTextCompression pdfTextCompression)
        {
            //ExStart
            //ExFor:PdfSaveOptions
            //ExFor:PdfSaveOptions.TextCompression
            //ExFor:PdfTextCompression
            //ExSummary:Shows how to apply text compression when saving a document to PDF.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            for (int i = 0; i < 100; i++)
                builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, " +
                                "sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Set the "TextCompression" property to "PdfTextCompression.None" to not apply any
            // compression to text when we save the document to PDF.
            // Set the "TextCompression" property to "PdfTextCompression.Flate" to apply ZIP compression
            // to text when we save the document to PDF. The larger the document, the bigger the impact that this will have.
            options.TextCompression = pdfTextCompression;

            doc.Save(ArtifactsDir + "PdfSaveOptions.TextCompression.pdf", options);
            //ExEnd

            switch (pdfTextCompression)
            {
                case PdfTextCompression.None:
                    Assert.That(60000,
                        Is.LessThan(new FileInfo(ArtifactsDir + "PdfSaveOptions.TextCompression.pdf").Length));
                    TestUtil.FileContainsString("5 0 obj\r\n<</Length 9 0 R>>stream",
                        ArtifactsDir + "PdfSaveOptions.TextCompression.pdf");
                    break;
                case PdfTextCompression.Flate:
                    Assert.That(30000,
                        Is.AtLeast(new FileInfo(ArtifactsDir + "PdfSaveOptions.TextCompression.pdf").Length));
                    TestUtil.FileContainsString("5 0 obj\r\n<</Length 9 0 R/Filter /FlateDecode>>stream",
                        ArtifactsDir + "PdfSaveOptions.TextCompression.pdf");
                    break;
            }
        }

        [TestCase(PdfImageCompression.Auto)]
        [TestCase(PdfImageCompression.Jpeg)]
        public void ImageCompression(PdfImageCompression pdfImageCompression)
        {
            //ExStart
            //ExFor:PdfSaveOptions.ImageCompression
            //ExFor:PdfSaveOptions.JpegQuality
            //ExFor:PdfImageCompression
            //ExSummary:Shows how to specify a compression type for all images in a document that we are converting to PDF.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Jpeg image:");
            builder.InsertImage(ImageDir + "Logo.jpg");
            builder.InsertParagraph();
            builder.Writeln("Png image:");
            builder.InsertImage(ImageDir + "Transparent background logo.png");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // Set the "ImageCompression" property to "PdfImageCompression.Auto" to use the
            // "ImageCompression" property to control the quality of the Jpeg images that end up in the output PDF.
            // Set the "ImageCompression" property to "PdfImageCompression.Jpeg" to use the
            // "ImageCompression" property to control the quality of all images that end up in the output PDF.
            pdfSaveOptions.ImageCompression = pdfImageCompression;

            // Set the "JpegQuality" property to "10" to strengthen compression at the cost of image quality.
            pdfSaveOptions.JpegQuality = 10;

            doc.Save(ArtifactsDir + "PdfSaveOptions.ImageCompression.pdf", pdfSaveOptions);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.ImageCompression.pdf");
            Stream pdfDocImageStream = pdfDocument.Pages[1].Resources.Images[1].ToStream();

            using (pdfDocImageStream)
            {
                TestUtil.VerifyImage(400, 400, pdfDocImageStream);
            }

            pdfDocImageStream = pdfDocument.Pages[1].Resources.Images[2].ToStream();

            using (pdfDocImageStream)
            {
                switch (pdfImageCompression)
                {
                    case PdfImageCompression.Auto:
                        Assert.That(50000, 
                            Is.LessThan(new FileInfo(ArtifactsDir + "PdfSaveOptions.ImageCompression.pdf").Length));
#if NET462
                        Assert.Throws<ArgumentException>(() => { TestUtil.VerifyImage(400, 400, pdfDocImageStream); });
#elif NETCOREAPP2_1
                        Assert.Throws<NullReferenceException>(() => { TestUtil.VerifyImage(400, 400, pdfDocImageStream); });
#endif
                        break;
                    case PdfImageCompression.Jpeg:
                        Assert.That(42000, 
                            Is.AtLeast(new FileInfo(ArtifactsDir + "PdfSaveOptions.ImageCompression.pdf").Length));
                        TestUtil.VerifyImage(400, 400, pdfDocImageStream);
                        break;
                }
            }
#endif
                }

        [TestCase(PdfImageColorSpaceExportMode.Auto)]
        [TestCase(PdfImageColorSpaceExportMode.SimpleCmyk)]
        public void ImageColorSpaceExportMode(PdfImageColorSpaceExportMode pdfImageColorSpaceExportMode)
        {
            //ExStart
            //ExFor:PdfImageColorSpaceExportMode
            //ExFor:PdfSaveOptions.ImageColorSpaceExportMode
            //ExSummary:Shows how to set a different color space for images in a document as we export it to PDF.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Jpeg image:");
            builder.InsertImage(ImageDir + "Logo.jpg");
            builder.InsertParagraph();
            builder.Writeln("Png image:");
            builder.InsertImage(ImageDir + "Transparent background logo.png");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // Set the "ImageColorSpaceExportMode" property to "PdfImageColorSpaceExportMode.Auto" to get Aspose.Words to
            // automatically select the color space for images in the document that it converts to PDF.
            // In most cases, the color space will be RGB.
            // Set the "ImageColorSpaceExportMode" property to "PdfImageColorSpaceExportMode.SimpleCmyk"
            // to use the CMYK color space for all images in the saved PDF.
            // Aspose.Words will also apply Flate compression to all images and ignore the "ImageCompression" property's value.
            pdfSaveOptions.ImageColorSpaceExportMode = pdfImageColorSpaceExportMode;

            doc.Save(ArtifactsDir + "PdfSaveOptions.ImageColorSpaceExportMode.pdf", pdfSaveOptions);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.ImageColorSpaceExportMode.pdf");
            XImage pdfDocImage = pdfDocument.Pages[1].Resources.Images[1];

            switch (pdfImageColorSpaceExportMode)
            {
                case PdfImageColorSpaceExportMode.Auto:
                    Assert.That(20000, Is.LessThan(pdfDocImage.ToStream().Length));
                    break;
                case PdfImageColorSpaceExportMode.SimpleCmyk:
                    Assert.That(100000, Is.LessThan(pdfDocImage.ToStream().Length));
                    break;
            }

            Assert.AreEqual(400, pdfDocImage.Width);
            Assert.AreEqual(400, pdfDocImage.Height);
            Assert.AreEqual(ColorType.Rgb, pdfDocImage.GetColorType());

            pdfDocImage = pdfDocument.Pages[1].Resources.Images[2];

            switch (pdfImageColorSpaceExportMode)
            {
                case PdfImageColorSpaceExportMode.Auto:
                    Assert.That(25000, Is.AtLeast(pdfDocImage.ToStream().Length));
                    break;
                case PdfImageColorSpaceExportMode.SimpleCmyk:
                    Assert.That(18000, Is.LessThan(pdfDocImage.ToStream().Length));
                    break;
            }

            Assert.AreEqual(400, pdfDocImage.Width);
            Assert.AreEqual(400, pdfDocImage.Height);
            Assert.AreEqual(ColorType.Rgb, pdfDocImage.GetColorType());
#endif
        }

        [Test]
        public void DownsampleOptions()
        {
            //ExStart
            //ExFor:DownsampleOptions
            //ExFor:DownsampleOptions.DownsampleImages
            //ExFor:DownsampleOptions.Resolution
            //ExFor:DownsampleOptions.ResolutionThreshold
            //ExFor:PdfSaveOptions.DownsampleOptions
            //ExSummary:Shows how to change the resolution of images in the PDF document.
            Document doc = new Document(MyDir + "Images.docx");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // By default, Aspose.Words downsample all images in a document that we save to PDF to 220 ppi.
            Assert.True(options.DownsampleOptions.DownsampleImages);
            Assert.AreEqual(220, options.DownsampleOptions.Resolution);
            Assert.AreEqual(0, options.DownsampleOptions.ResolutionThreshold);

            doc.Save(ArtifactsDir + "PdfSaveOptions.DownsampleOptions.Default.pdf", options);

            // Set the "Resolution" property to "36" to downsample all images to 36 ppi.
            options.DownsampleOptions.Resolution = 36;

            // Set the "ResolutionThreshold" property to only apply the downsampling to
            // images with a resolution that is above 128 ppi.
            options.DownsampleOptions.ResolutionThreshold = 128;

            // Only the first two images from the document will be downsampled at this stage.
            doc.Save(ArtifactsDir + "PdfSaveOptions.DownsampleOptions.LowerResolution.pdf", options);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.DownsampleOptions.Default.pdf");
            XImage pdfDocImage = pdfDocument.Pages[1].Resources.Images[1];

            Assert.That(300000, Is.LessThan(pdfDocImage.ToStream().Length));
            Assert.AreEqual(ColorType.Rgb, pdfDocImage.GetColorType());
#endif
        }

        [TestCase(ColorMode.Grayscale)]
        [TestCase(ColorMode.Normal)]
        public void ColorRendering(ColorMode colorMode)
        {
            //ExStart
            //ExFor:PdfSaveOptions
            //ExFor:ColorMode
            //ExFor:FixedPageSaveOptions.ColorMode
            //ExSummary:Shows how to change image color with saving options property.
            Document doc = new Document(MyDir + "Images.docx");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            // Set the "ColorMode" property to "Grayscale" to render all images from the document in black and white.
            // The size of the output document may be larger with this setting.
            // Set the "ColorMode" property to "Normal" to render all images in color.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { ColorMode = colorMode };
            
            doc.Save(ArtifactsDir + "PdfSaveOptions.ColorRendering.pdf", pdfSaveOptions);
            //ExEnd

            #if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.ColorRendering.pdf");
            XImage pdfDocImage = pdfDocument.Pages[1].Resources.Images[1];

            switch (colorMode)
            {
                case ColorMode.Normal:
                    Assert.That(300000, Is.LessThan(pdfDocImage.ToStream().Length));
                    Assert.AreEqual(ColorType.Rgb, pdfDocImage.GetColorType());
                    break;
                case ColorMode.Grayscale:
                    Assert.That(1000000, Is.LessThan(pdfDocImage.ToStream().Length));
                    Assert.AreEqual(ColorType.Grayscale, pdfDocImage.GetColorType());
                    break;
            }
            #endif
        }

        [TestCase(false)]
        [TestCase(true)]
        public void DocTitle(bool displayDocTitle)
        {
            //ExStart
            //ExFor:PdfSaveOptions.DisplayDocTitle
            //ExSummary:Shows how to display the title of the document as the title bar.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            doc.BuiltInDocumentProperties.Title = "Windows bar pdf title";

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            // Set the "DisplayDocTitle" to "true" to get some PDF readers, such as Adobe Acrobat Pro,
            // to display the value of the document's "Title" built-in property in the tab that belongs to this document.
            // Set the "DisplayDocTitle" to "false" to get such readers to display the document's filename.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions { DisplayDocTitle = displayDocTitle };

            doc.Save(ArtifactsDir + "PdfSaveOptions.DocTitle.pdf", pdfSaveOptions);
            //ExEnd

            #if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.DocTitle.pdf");

            Assert.AreEqual(displayDocTitle, pdfDocument.DisplayDocTitle);
            Assert.AreEqual("Windows bar pdf title", pdfDocument.Info.Title);
            #endif
        }

        [TestCase(false)]
        [TestCase(true)]
        public void MemoryOptimization(bool memoryOptimization)
        {
            //ExStart
            //ExFor:SaveOptions.CreateSaveOptions(SaveFormat)
            //ExFor:SaveOptions.MemoryOptimization
            //ExSummary:Shows an option to optimize memory consumption when rendering large documents to PDF.
            Document doc = new Document(MyDir + "Rendering.docx");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);

            // Set the "MemoryOptimization" property to "true" to lower the memory footprint of large documents' saving operations
            // at the cost of increasing the duration of the operation.
            // Set the "MemoryOptimization" property to "false" to save the document as a PDF normally.
            saveOptions.MemoryOptimization = memoryOptimization;

            doc.Save(ArtifactsDir + "PdfSaveOptions.MemoryOptimization.pdf", saveOptions);
            //ExEnd
        }

        [TestCase(@"https://www.google.com/search?q= aspose", "https://www.google.com/search?q=%20aspose", true)]
        [TestCase(@"https://www.google.com/search?q=%20aspose", "https://www.google.com/search?q=%20aspose", true)]
        [TestCase(@"https://www.google.com/search?q= aspose", "https://www.google.com/search?q= aspose", false)]
        [TestCase(@"https://www.google.com/search?q=%20aspose", "https://www.google.com/search?q=%20aspose", false)]
        public void EscapeUri(string uri, string result, bool isEscaped)
        {
            //ExStart
            //ExFor:PdfSaveOptions.EscapeUri
            //ExSummary:Shows how to escape hyperlinks in the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertHyperlink("Testlink", uri, false);

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Set the "EscapeUri" property to "true" if links in the document contain characters,
            // such as the blank space, that we need to replace with escape sequences, such as "%20".
            // Set the "EscapeUri" property to "false" if we are sure that this document's links
            // do not need any such escape character substitution.
            options.EscapeUri = isEscaped;

            doc.Save(ArtifactsDir + "PdfSaveOptions.EscapedUri.pdf", options);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.EscapedUri.pdf");

            Page page = pdfDocument.Pages[1];
            LinkAnnotation linkAnnot = (LinkAnnotation)page.Annotations[1];

            GoToURIAction action = (GoToURIAction)linkAnnot.Action;

            Assert.AreEqual(result, action.URI);
#endif
        }

        [TestCase(false)]
        [TestCase(true)]
        public void OpenHyperlinksInNewWindow(bool openHyperlinksInNewWindow)
        {
            //ExStart
            //ExFor:PdfSaveOptions.OpenHyperlinksInNewWindow
            //ExSummary:Shows how to save hyperlinks in a document we convert to PDF so that they open new pages when we click on them.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertHyperlink("Testlink", @"https://www.google.com/search?q=%20aspose", false);

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Set the "OpenHyperlinksInNewWindow" property to "true" to save all hyperlinks using Javascript code
            // that forces readers to open these links in new windows/browser tabs.
            // Set the "OpenHyperlinksInNewWindow" property to "false" to save all hyperlinks normally.
            options.OpenHyperlinksInNewWindow = openHyperlinksInNewWindow;

            doc.Save(ArtifactsDir + "PdfSaveOptions.OpenHyperlinksInNewWindow.pdf", options);
            //ExEnd

            if (openHyperlinksInNewWindow)
                TestUtil.FileContainsString(
                    "<</Type /Annot/Subtype /Link/Rect [70.84999847 707.35101318 110.17799377 721.15002441]/BS " +
                    "<</Type/Border/S/S/W 0>>/A<</Type /Action/S /JavaScript/JS(app.launchURL\\(\"https://www.google.com/search?q=%20aspose\", true\\);)>>>>",
                    ArtifactsDir + "PdfSaveOptions.OpenHyperlinksInNewWindow.pdf");
            else
                TestUtil.FileContainsString(
                    "<</Type /Annot/Subtype /Link/Rect [70.84999847 707.35101318 110.17799377 721.15002441]/BS " +
                    "<</Type/Border/S/S/W 0>>/A<</Type /Action/S /URI/URI(https://www.google.com/search?q=%20aspose)>>>>",
                    ArtifactsDir + "PdfSaveOptions.OpenHyperlinksInNewWindow.pdf");

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument =
                new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.OpenHyperlinksInNewWindow.pdf");

            Page page = pdfDocument.Pages[1];
            LinkAnnotation linkAnnot = (LinkAnnotation) page.Annotations[1];

            Assert.AreEqual(openHyperlinksInNewWindow ? typeof(JavascriptAction) : typeof(GoToURIAction),
                linkAnnot.Action.GetType());
#endif
        }

        //ExStart
        //ExFor:MetafileRenderingMode
        //ExFor:MetafileRenderingOptions
        //ExFor:MetafileRenderingOptions.EmulateRasterOperations
        //ExFor:MetafileRenderingOptions.RenderingMode
        //ExFor:IWarningCallback
        //ExFor:FixedPageSaveOptions.MetafileRenderingOptions
        //ExSummary:Shows added a fallback to bitmap rendering and changing type of warnings about unsupported metafile records.
        [Test, Category("SkipMono")] //ExSkip
        public void HandleBinaryRasterWarnings()
        {
            Document doc = new Document(MyDir + "WMF with image.docx");

            MetafileRenderingOptions metafileRenderingOptions = new MetafileRenderingOptions();

            // Set the "EmulateRasterOperations" property to "false" to fall back to bitmap when
            // it encounters a metafile, which will require raster operations to render in the output PDF.
            metafileRenderingOptions.EmulateRasterOperations = false;

            // Set the "RenderingMode" property to "VectorWithFallback" to try to render every metafile using vector graphics.
            metafileRenderingOptions.RenderingMode = MetafileRenderingMode.VectorWithFallback;

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF and applies the configuration
            // in our MetafileRenderingOptions object to the saving operation.
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.MetafileRenderingOptions = metafileRenderingOptions;

            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.WarningCallback = callback;

            doc.Save(ArtifactsDir + "PdfSaveOptions.HandleBinaryRasterWarnings.pdf", saveOptions);

            Assert.AreEqual(1, callback.Warnings.Count);
            Assert.AreEqual("'R2_XORPEN' binary raster operation is partly supported.",
                callback.Warnings[0].Description);
        }

        /// <summary>
        /// Prints and collects formatting loss-related warnings that occur upon saving a document.
        /// </summary>
        public class HandleDocumentWarnings : IWarningCallback
        {
            public void Warning(WarningInfo info)
            {
                if (info.WarningType == WarningType.MinorFormattingLoss)
                {
                    Console.WriteLine("Unsupported operation: " + info.Description);
                    Warnings.Warning(info);
                }
            }

            public WarningInfoCollection Warnings = new WarningInfoCollection();
        }
        //ExEnd

        [TestCase(Aspose.Words.Saving.HeaderFooterBookmarksExportMode.None)]
        [TestCase(Aspose.Words.Saving.HeaderFooterBookmarksExportMode.First)]
        [TestCase(Aspose.Words.Saving.HeaderFooterBookmarksExportMode.All)]
        public void HeaderFooterBookmarksExportMode(HeaderFooterBookmarksExportMode headerFooterBookmarksExportMode)
        {
            //ExStart
            //ExFor:HeaderFooterBookmarksExportMode
            //ExFor:OutlineOptions
            //ExFor:OutlineOptions.DefaultBookmarksOutlineLevel
            //ExFor:PdfSaveOptions.HeaderFooterBookmarksExportMode
            //ExFor:PdfSaveOptions.PageMode
            //ExFor:PdfPageMode
            //ExSummary:Shows to process bookmarks in headers/footers in a document that we are rendering to PDF.
            Document doc = new Document(MyDir + "Bookmarks in headers and footers.docx");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Set the "PageMode" property to "PdfPageMode.UseOutlines" to display the outline navigation pane in the output PDF.
            saveOptions.PageMode = PdfPageMode.UseOutlines;

            // Set the "DefaultBookmarksOutlineLevel" property to "1" to display all
            // bookmarks at the first level of the outline in the output PDF.
            saveOptions.OutlineOptions.DefaultBookmarksOutlineLevel = 1;

            // Set the "HeaderFooterBookmarksExportMode" property to "HeaderFooterBookmarksExportMode.None" to
            // not export any bookmarks that are inside headers/footers.
            // Set the "HeaderFooterBookmarksExportMode" property to "HeaderFooterBookmarksExportMode.First" to
            // only export bookmarks in the first section's header/footers.
            // Set the "HeaderFooterBookmarksExportMode" property to "HeaderFooterBookmarksExportMode.All" to
            // export bookmarks that are in all headers/footers.
            saveOptions.HeaderFooterBookmarksExportMode = headerFooterBookmarksExportMode;

            doc.Save(ArtifactsDir + "PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf", saveOptions);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDoc =
                new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf");
            string inputDocLocaleName = new CultureInfo(doc.Styles.DefaultFont.LocaleId).Name;

            TextFragmentAbsorber textFragmentAbsorber = new TextFragmentAbsorber();
            pdfDoc.Pages.Accept(textFragmentAbsorber);
            switch (headerFooterBookmarksExportMode)
            {
                case Aspose.Words.Saving.HeaderFooterBookmarksExportMode.None:
                    TestUtil.FileContainsString($"<</Type /Catalog/Pages 3 0 R/Lang({inputDocLocaleName})>>\r\n",
                        ArtifactsDir + "PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf");

                    Assert.AreEqual(0, pdfDoc.Outlines.Count);
                    break;
                case Aspose.Words.Saving.HeaderFooterBookmarksExportMode.First:
                case Aspose.Words.Saving.HeaderFooterBookmarksExportMode.All:
                    TestUtil.FileContainsString(
                        $"<</Type /Catalog/Pages 3 0 R/Outlines 13 0 R/PageMode /UseOutlines/Lang({inputDocLocaleName})>>",
                        ArtifactsDir + "PdfSaveOptions.HeaderFooterBookmarksExportMode.pdf");

                    OutlineCollection outlineItemCollection = pdfDoc.Outlines;

                    Assert.AreEqual(4, outlineItemCollection.Count);
                    Assert.AreEqual("Bookmark_1", outlineItemCollection[1].Title);
                    Assert.AreEqual("1 XYZ 233 806 0", outlineItemCollection[1].Destination.ToString());

                    Assert.AreEqual("Bookmark_2", outlineItemCollection[2].Title);
                    Assert.AreEqual("1 XYZ 84 47 0", outlineItemCollection[2].Destination.ToString());

                    Assert.AreEqual("Bookmark_3", outlineItemCollection[3].Title);
                    Assert.AreEqual("2 XYZ 85 806 0", outlineItemCollection[3].Destination.ToString());

                    Assert.AreEqual("Bookmark_4", outlineItemCollection[4].Title);
                    Assert.AreEqual("2 XYZ 85 48 0", outlineItemCollection[4].Destination.ToString());
                    break;
            }
#endif
        }

        [Test]
        public void UnsupportedImageFormatWarning()
        {
            Document doc = new Document(MyDir + "Corrupted image.docx");

            SaveWarningCallback saveWarningCallback = new SaveWarningCallback();
            doc.WarningCallback = saveWarningCallback;

            doc.Save(ArtifactsDir + "PdfSaveOption.UnsupportedImageFormatWarning.pdf", SaveFormat.Pdf);

            Assert.That(saveWarningCallback.SaveWarnings[0].Description,
                Is.EqualTo("Image can not be processed. Possibly unsupported image format."));
        }

        public class SaveWarningCallback : IWarningCallback
        {
            public void Warning(WarningInfo info)
            {
                if (info.WarningType == WarningType.MinorFormattingLoss)
                {
                    Console.WriteLine($"{info.WarningType}: {info.Description}.");
                    SaveWarnings.Warning(info);
                }
            }

            internal WarningInfoCollection SaveWarnings = new WarningInfoCollection();
		}
		
		[TestCase(false)]
        [TestCase(true)]
        public void FontsScaledToMetafileSize(bool scaleWmfFonts)
        {
            //ExStart
            //ExFor:MetafileRenderingOptions.ScaleWmfFontsToMetafileSize
            //ExSummary:Shows how to WMF fonts scaling according to metafile size on the page.
            Document doc = new Document(MyDir + "WMF with text.docx");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Set the "ScaleWmfFontsToMetafileSize" property to "true" to scale fonts
            // that format text within WMF images according to the size of the metafile on the page.
            // Set the "ScaleWmfFontsToMetafileSize" property to "false" to
            // preserve the default scale of these fonts.
            saveOptions.MetafileRenderingOptions.ScaleWmfFontsToMetafileSize = scaleWmfFonts;

            doc.Save(ArtifactsDir + "PdfSaveOptions.FontsScaledToMetafileSize.pdf", saveOptions);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.FontsScaledToMetafileSize.pdf");
            TextFragmentAbsorber textAbsorber = new TextFragmentAbsorber();

            pdfDocument.Pages[1].Accept(textAbsorber);
            Rectangle textFragmentRectangle = textAbsorber.TextFragments[3].Rectangle;

            Assert.AreEqual(scaleWmfFonts ? 1.589d : 5.045d, textFragmentRectangle.Width, 0.001d);
#endif
        }

        [TestCase(false)]
        [TestCase(true)]
        public void EmbedFullFonts(bool embedFullFonts)
        {
            //ExStart
            //ExFor:PdfSaveOptions.#ctor
            //ExFor:PdfSaveOptions.EmbedFullFonts
            //ExSummary:Shows how to enable or disable subsetting when embedding fonts while rendering a document to PDF.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Arial";
            builder.Writeln("Hello world!");
            builder.Font.Name = "Arvo";
            builder.Writeln("The quick brown fox jumps over the lazy dog.");

            // Configure our font sources to ensure that we have access to both the fonts in this document.
            FontSourceBase[] originalFontsSources = FontSettings.DefaultInstance.GetFontsSources();
            Aspose.Words.Fonts.FolderFontSource folderFontSource = new Aspose.Words.Fonts.FolderFontSource(FontsDir, true);
            FontSettings.DefaultInstance.SetFontsSources(new[] { originalFontsSources[0], folderFontSource });

            FontSourceBase[] fontSources = FontSettings.DefaultInstance.GetFontsSources();
            Assert.True(fontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arial"));
            Assert.True(fontSources[1].GetAvailableFonts().Any(f => f.FullFontName == "Arvo"));

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Since our document contains a custom font, embedding in the output document may be desirable.
            // Set the "EmbedFullFonts" property to "true" to embed every glyph of every embedded font in the output PDF.
            // The document's size may become very large, but we will have full use of all fonts if we edit the PDF.
            // Set the "EmbedFullFonts" property to "false" to apply subsetting to fonts, saving only the glyphs
            // that the document is using. The file will be considerably smaller,
            // but we may need access to any custom fonts if we edit the document.
            options.EmbedFullFonts = embedFullFonts;

            doc.Save(ArtifactsDir + "PdfSaveOptions.EmbedFullFonts.pdf", options);

            if (embedFullFonts)
                Assert.That(500000, Is.LessThan(new FileInfo(ArtifactsDir + "PdfSaveOptions.EmbedFullFonts.pdf").Length));
            else
                Assert.That(25000, Is.AtLeast(new FileInfo(ArtifactsDir + "PdfSaveOptions.EmbedFullFonts.pdf").Length));

            // Restore the original font sources.
            FontSettings.DefaultInstance.SetFontsSources(originalFontsSources);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.EmbedFullFonts.pdf");

            Aspose.Pdf.Text.Font[] pdfDocFonts = pdfDocument.FontUtilities.GetAllFonts();

            Assert.AreEqual("ArialMT", pdfDocFonts[0].FontName);
            Assert.AreNotEqual(embedFullFonts, pdfDocFonts[0].IsSubset);

            Assert.AreEqual("Arvo", pdfDocFonts[1].FontName);
            Assert.AreNotEqual(embedFullFonts, pdfDocFonts[1].IsSubset);
#endif
        }

        [TestCase(PdfFontEmbeddingMode.EmbedAll)]
        [TestCase(PdfFontEmbeddingMode.EmbedNone)]
        [TestCase(PdfFontEmbeddingMode.EmbedNonstandard)]
        public void EmbedWindowsFonts(PdfFontEmbeddingMode pdfFontEmbeddingMode)
        {
            //ExStart
            //ExFor:PdfSaveOptions.FontEmbeddingMode
            //ExFor:PdfFontEmbeddingMode
            //ExSummary:Shows how to set Aspose.Words to skip embedding Arial and Times New Roman fonts into a PDF document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // "Arial" is a standard font, and "Courier New" is a nonstandard font.
            builder.Font.Name = "Arial";
            builder.Writeln("Hello world!");
            builder.Font.Name = "Courier New";
            builder.Writeln("The quick brown fox jumps over the lazy dog.");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Set the "EmbedFullFonts" property to "true" to embed every glyph of every embedded font in the output PDF.
            options.EmbedFullFonts = true;

            // Set the "FontEmbeddingMode" property to "EmbedAll" to embed all fonts in the output PDF.
            // Set the "FontEmbeddingMode" property to "EmbedNonstandard" to only allow nonstandard fonts' embedding in the output PDF.
            // Set the "FontEmbeddingMode" property to "EmbedNone" to not embed any fonts in the output PDF.
            options.FontEmbeddingMode = pdfFontEmbeddingMode;

            doc.Save(ArtifactsDir + "PdfSaveOptions.EmbedWindowsFonts.pdf", options);

            switch (pdfFontEmbeddingMode)
            {
                case PdfFontEmbeddingMode.EmbedAll:
                    Assert.That(1000000, Is.LessThan(new FileInfo(ArtifactsDir + "PdfSaveOptions.EmbedWindowsFonts.pdf").Length));
                    break;
                case PdfFontEmbeddingMode.EmbedNonstandard:
                    Assert.That(480000, Is.LessThan(new FileInfo(ArtifactsDir + "PdfSaveOptions.EmbedWindowsFonts.pdf").Length));
                    break;
                case PdfFontEmbeddingMode.EmbedNone:
                    Assert.That(4000, Is.AtLeast(new FileInfo(ArtifactsDir + "PdfSaveOptions.EmbedWindowsFonts.pdf").Length));
                    break;
            }
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.EmbedWindowsFonts.pdf");

            Aspose.Pdf.Text.Font[] pdfDocFonts = pdfDocument.FontUtilities.GetAllFonts();

            Assert.AreEqual("ArialMT", pdfDocFonts[0].FontName);
            Assert.AreEqual(pdfFontEmbeddingMode == PdfFontEmbeddingMode.EmbedAll, 
                pdfDocFonts[0].IsEmbedded);

            Assert.AreEqual("CourierNewPSMT", pdfDocFonts[1].FontName);
            Assert.AreEqual(pdfFontEmbeddingMode == PdfFontEmbeddingMode.EmbedAll || pdfFontEmbeddingMode == PdfFontEmbeddingMode.EmbedNonstandard, 
                pdfDocFonts[1].IsEmbedded);
#endif
        }

        [TestCase(false)]
        [TestCase(true)]
        public void EmbedCoreFonts(bool useCoreFonts)
        {
            //ExStart
            //ExFor:PdfSaveOptions.UseCoreFonts
            //ExSummary:Shows how enable/disable PDF Type 1 font substitution.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Arial";
            builder.Writeln("Hello world!");
            builder.Font.Name = "Courier New";
            builder.Writeln("The quick brown fox jumps over the lazy dog.");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Set the "UseCoreFonts" property to "true" to replace some fonts,
            // including the two fonts in our document, with their PDF Type 1 equivalents.
            // Set the "UseCoreFonts" property to "false" to not apply PDF Type 1 fonts.
            options.UseCoreFonts = useCoreFonts;

            doc.Save(ArtifactsDir + "PdfSaveOptions.EmbedCoreFonts.pdf", options);

            if (useCoreFonts)
                Assert.That(3000, Is.AtLeast(new FileInfo(ArtifactsDir + "PdfSaveOptions.EmbedCoreFonts.pdf").Length));
            else
                Assert.That(30000, Is.LessThan(new FileInfo(ArtifactsDir + "PdfSaveOptions.EmbedCoreFonts.pdf").Length));
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.EmbedCoreFonts.pdf");

            Aspose.Pdf.Text.Font[] pdfDocFonts = pdfDocument.FontUtilities.GetAllFonts();

            if (useCoreFonts)
            {
                Assert.AreEqual("Helvetica", pdfDocFonts[0].FontName);
                Assert.AreEqual("Courier", pdfDocFonts[1].FontName);
            }
            else
            {
                Assert.AreEqual("ArialMT", pdfDocFonts[0].FontName);
                Assert.AreEqual("CourierNewPSMT", pdfDocFonts[1].FontName);
            }

            Assert.AreNotEqual(useCoreFonts, pdfDocFonts[0].IsEmbedded);
            Assert.AreNotEqual(useCoreFonts, pdfDocFonts[1].IsEmbedded);
#endif
        }

        [TestCase(false)]
        [TestCase(true)]
        public void AdditionalTextPositioning(bool applyAdditionalTextPositioning)
        {
            //ExStart
            //ExFor:PdfSaveOptions.AdditionalTextPositioning
            //ExSummary:Show how to write additional text positioning operators.
            Document doc = new Document(MyDir + "Text positioning operators.docx");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                TextCompression = PdfTextCompression.None,

                // Set the "AdditionalTextPositioning" property to "true" to attempt to fix incorrect
                // element positioning in the output PDF, should there be any, at the cost of increased file size.
                // Set the "AdditionalTextPositioning" property to "false" to render the document as usual.
                AdditionalTextPositioning = applyAdditionalTextPositioning
            };

            doc.Save(ArtifactsDir + "PdfSaveOptions.AdditionalTextPositioning.pdf", saveOptions);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument =
                new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.AdditionalTextPositioning.pdf");
            TextFragmentAbsorber textAbsorber = new TextFragmentAbsorber();

            pdfDocument.Pages[1].Accept(textAbsorber);

            SetGlyphsPositionShowText tjOperator =
                (SetGlyphsPositionShowText) textAbsorber.TextFragments[1].Page.Contents[85];

            if (applyAdditionalTextPositioning)
            {
                Assert.That(100000,
                    Is.LessThan(new FileInfo(ArtifactsDir + "PdfSaveOptions.AdditionalTextPositioning.pdf").Length));
                Assert.AreEqual(
                    "[0 (S) 0 (a) 0 (m) 0 (s) 0 (t) 0 (a) -1 (g) 1 (,) 0 ( ) 0 (1) 0 (0) 0 (.) 0 ( ) 0 (N) 0 (o) 0 (v) 0 (e) 0 (m) 0 (b) 0 (e) 0 (r) -1 ( ) 1 (2) -1 (0) 0 (1) 0 (8)] TJ",
                    tjOperator.ToString());
            }
            else
            {
                Assert.That(97000,
                    Is.LessThan(new FileInfo(ArtifactsDir + "PdfSaveOptions.AdditionalTextPositioning.pdf").Length));
                Assert.AreEqual("[(Samsta) -1 (g) 1 (, 10. November) -1 ( ) 1 (2) -1 (018)] TJ", tjOperator.ToString());
            }
#endif
        }

        [TestCase(false, Category = "SkipMono")]
        [TestCase(true, Category = "SkipMono")]
        public void SaveAsPdfBookFold(bool renderTextAsBookfold)
        {
            //ExStart
            //ExFor:PdfSaveOptions.UseBookFoldPrintingSettings
            //ExSummary:Shows how to save a document to the PDF format in the form of a book fold.
            Document doc = new Document(MyDir + "Paragraphs.docx");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Set the "UseBookFoldPrintingSettings" property to "true" to arrange the contents
            // in the output PDF in a way that helps us use it to make a booklet.
            // Set the "UseBookFoldPrintingSettings" property to "false" to render the PDF normally.
            options.UseBookFoldPrintingSettings = renderTextAsBookfold;

            // If we are rendering the document as a booklet, we must set the "MultiplePages"
            // properties of the page setup objects of all sections to "MultiplePagesType.BookFoldPrinting".
            if (renderTextAsBookfold)
                foreach (Section s in doc.Sections)
                {
                    s.PageSetup.MultiplePages = MultiplePagesType.BookFoldPrinting;
                }

            // Once we print this document on both sides of the pages, we can fold all the pages down the middle at once,
            // and the contents will line up in a way that creates a booklet.
            doc.Save(ArtifactsDir + "PdfSaveOptions.SaveAsPdfBookFold.pdf", options);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.SaveAsPdfBookFold.pdf");
            TextAbsorber textAbsorber = new TextAbsorber();

            pdfDocument.Pages.Accept(textAbsorber);

            if (renderTextAsBookfold)
            {
                Assert.True(textAbsorber.Text.IndexOf("Heading #1", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #2", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #2", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #3", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #3", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #4", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #4", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #5", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #5", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #6", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #6", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #7", StringComparison.Ordinal));
                Assert.False(textAbsorber.Text.IndexOf("Heading #7", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #8", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #8", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #9", StringComparison.Ordinal));
                Assert.False(textAbsorber.Text.IndexOf("Heading #9", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #10", StringComparison.Ordinal));
            }
            else
            {
                Assert.True(textAbsorber.Text.IndexOf("Heading #1", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #2", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #2", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #3", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #3", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #4", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #4", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #5", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #5", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #6", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #6", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #7", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #7", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #8", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #8", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #9", StringComparison.Ordinal));
                Assert.True(textAbsorber.Text.IndexOf("Heading #9", StringComparison.Ordinal) < textAbsorber.Text.IndexOf("Heading #10", StringComparison.Ordinal));
            }
#endif
        }

        [Test]
        public void ZoomBehaviour()
        {
            //ExStart
            //ExFor:PdfSaveOptions.ZoomBehavior
            //ExFor:PdfSaveOptions.ZoomFactor
            //ExFor:PdfZoomBehavior
            //ExSummary:Shows how to set the default zooming that a reader applies when opening a rendered PDF document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            // Set the "ZoomBehavior" property to "PdfZoomBehavior.ZoomFactor" to get a PDF reader to
            // apply a percentage-based zoom factor when we open the document with it.
            // Set the "ZoomFactor" property to "25" to give the zoom factor a value of 25%.
            PdfSaveOptions options = new PdfSaveOptions
            {
                ZoomBehavior = PdfZoomBehavior.ZoomFactor,
                ZoomFactor = 25
            };

            // When we open this document using a reader such as Adobe Acrobat, we will see the document scaled at 1/4 of its actual size.
            doc.Save(ArtifactsDir + "PdfSaveOptions.ZoomBehaviour.pdf", options);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.ZoomBehaviour.pdf");
            GoToAction action = (GoToAction)pdfDocument.OpenAction;

            Assert.AreEqual(0.25d, (action.Destination as XYZExplicitDestination).Zoom);
#endif
        }

        [TestCase(PdfPageMode.FullScreen)]
        [TestCase(PdfPageMode.UseThumbs)]
        [TestCase(PdfPageMode.UseOC)]
        [TestCase(PdfPageMode.UseOutlines)]
        [TestCase(PdfPageMode.UseNone)]
        public void PageMode(PdfPageMode pageMode)
        {
            //ExStart
            //ExFor:PdfSaveOptions.PageMode
            //ExFor:PdfPageMode
            //ExSummary:Shows how to set instructions for some PDF readers to follow when opening an output document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Set the "PageMode" property to "PdfPageMode.FullScreen" to get the PDF reader to open the saved
            // document in full-screen mode, which takes over the monitor's display and has no controls visible.
            // Set the "PageMode" property to "PdfPageMode.UseThumbs" to get the PDF reader to display a separate panel
            // with a thumbnail for each page in the document.
            // Set the "PageMode" property to "PdfPageMode.UseOC" to get the PDF reader to display a separate panel
            // that allows us to work with any layers present in the document.
            // Set the "PageMode" property to "PdfPageMode.UseOutlines" to get the PDF reader
            // also to display the outline, if possible.
            // Set the "PageMode" property to "PdfPageMode.UseNone" to get the PDF reader to display just the document itself.
            options.PageMode = pageMode;

            doc.Save(ArtifactsDir + "PdfSaveOptions.PageMode.pdf", options);
            //ExEnd
            
            string docLocaleName = new CultureInfo(doc.Styles.DefaultFont.LocaleId).Name;

            switch (pageMode)
            {
                case PdfPageMode.FullScreen:
                    TestUtil.FileContainsString(
                        $"<</Type /Catalog/Pages 3 0 R/PageMode /FullScreen/Lang({docLocaleName})>>\r\n",
                        ArtifactsDir + "PdfSaveOptions.PageMode.pdf");
                    break;
                case PdfPageMode.UseThumbs:
                    TestUtil.FileContainsString(
                        $"<</Type /Catalog/Pages 3 0 R/PageMode /UseThumbs/Lang({docLocaleName})>>",
                        ArtifactsDir + "PdfSaveOptions.PageMode.pdf");
                    break;
                case PdfPageMode.UseOC:
                    TestUtil.FileContainsString(
                        $"<</Type /Catalog/Pages 3 0 R/PageMode /UseOC/Lang({docLocaleName})>>\r\n",
                        ArtifactsDir + "PdfSaveOptions.PageMode.pdf");
                    break;
                case PdfPageMode.UseOutlines:
                case PdfPageMode.UseNone:
                    TestUtil.FileContainsString($"<</Type /Catalog/Pages 3 0 R/Lang({docLocaleName})>>\r\n",
                        ArtifactsDir + "PdfSaveOptions.PageMode.pdf");
                    break;
            }

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.PageMode.pdf");

            switch (pageMode)
            {
                case PdfPageMode.UseNone:
                case PdfPageMode.UseOutlines:
                    Assert.AreEqual(Aspose.Pdf.PageMode.UseNone, pdfDocument.PageMode);
                    break;
                case PdfPageMode.UseThumbs:
                    Assert.AreEqual(Aspose.Pdf.PageMode.UseThumbs, pdfDocument.PageMode);
                    break;
                case PdfPageMode.FullScreen:
                    Assert.AreEqual(Aspose.Pdf.PageMode.FullScreen, pdfDocument.PageMode);
                    break;
                case PdfPageMode.UseOC:
                    Assert.AreEqual(Aspose.Pdf.PageMode.UseOC, pdfDocument.PageMode);
                    break;
            }
#endif
        }

        [TestCase(false)]
        [TestCase(true)]
        public void NoteHyperlinks(bool createNoteHyperlinks)
        {
            //ExStart
            //ExFor:PdfSaveOptions.CreateNoteHyperlinks
            //ExSummary:Shows how to make footnotes and endnotes function as hyperlinks.
            Document doc = new Document(MyDir + "Footnotes and endnotes.docx");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Set the "CreateNoteHyperlinks" property to "true" to turn all footnote/endnote symbols
            // in the text act as links that, upon clicking, take us to their respective footnotes/endnotes.
            // Set the "CreateNoteHyperlinks" property to "false" not to have footnote/endnote symbols link to anything.
            options.CreateNoteHyperlinks = createNoteHyperlinks;

            doc.Save(ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf", options);
            //ExEnd

            if (createNoteHyperlinks)
            {
                TestUtil.FileContainsString(
                    "<</Type /Annot/Subtype /Link/Rect [157.80099487 720.90106201 159.35600281 733.55004883]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 85 677 0]>>",
                    ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
                TestUtil.FileContainsString(
                    "<</Type /Annot/Subtype /Link/Rect [202.16900635 720.90106201 206.06201172 733.55004883]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 85 79 0]>>",
                    ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
                TestUtil.FileContainsString(
                    "<</Type /Annot/Subtype /Link/Rect [212.23199463 699.2510376 215.34199524 711.90002441]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 85 654 0]>>",
                    ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
                TestUtil.FileContainsString(
                    "<</Type /Annot/Subtype /Link/Rect [258.15499878 699.2510376 262.04800415 711.90002441]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 85 68 0]>>",
                    ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
                TestUtil.FileContainsString(
                    "<</Type /Annot/Subtype /Link/Rect [85.05000305 68.19905853 88.66500092 79.69805908]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 202 733 0]>>",
                    ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
                TestUtil.FileContainsString(
                    "<</Type /Annot/Subtype /Link/Rect [85.05000305 56.70005798 88.66500092 68.19905853]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 258 711 0]>>",
                    ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
                TestUtil.FileContainsString(
                    "<</Type /Annot/Subtype /Link/Rect [85.05000305 666.10205078 86.4940033 677.60107422]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 157 733 0]>>",
                    ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
                TestUtil.FileContainsString(
                    "<</Type /Annot/Subtype /Link/Rect [85.05000305 643.10406494 87.93800354 654.60308838]/BS <</Type/Border/S/S/W 0>>/Dest[4 0 R /XYZ 212 711 0]>>",
                    ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
            }
            else
            {
                if (!IsRunningOnMono())
                    Assert.Throws<AssertionException>(() =>
                        TestUtil.FileContainsString("<</Type /Annot/Subtype /Link/Rect",
                            ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf"));
            }

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.NoteHyperlinks.pdf");
            Page page = pdfDocument.Pages[1];
            AnnotationSelector annotationSelector = new AnnotationSelector(new LinkAnnotation(page, Rectangle.Trivial));

            page.Accept(annotationSelector);

            List<LinkAnnotation> linkAnnotations = annotationSelector.Selected.Cast<LinkAnnotation>().ToList();

            if (createNoteHyperlinks)
            {
                Assert.AreEqual(8, linkAnnotations.Count(a => a.AnnotationType == AnnotationType.Link));

                Assert.AreEqual("1 XYZ 85 677 0", linkAnnotations[0].Destination.ToString());
                Assert.AreEqual("1 XYZ 85 79 0", linkAnnotations[1].Destination.ToString());
                Assert.AreEqual("1 XYZ 85 654 0", linkAnnotations[2].Destination.ToString());
                Assert.AreEqual("1 XYZ 85 68 0", linkAnnotations[3].Destination.ToString());
                Assert.AreEqual("1 XYZ 202 733 0", linkAnnotations[4].Destination.ToString());
                Assert.AreEqual("1 XYZ 258 711 0", linkAnnotations[5].Destination.ToString());
                Assert.AreEqual("1 XYZ 157 733 0", linkAnnotations[6].Destination.ToString());
                Assert.AreEqual("1 XYZ 212 711 0", linkAnnotations[7].Destination.ToString());
            }
            else
            {
                Assert.AreEqual(0, annotationSelector.Selected.Count);
            }
#endif
        }

        [TestCase(PdfCustomPropertiesExport.None)]
        [TestCase(PdfCustomPropertiesExport.Standard)]
        [TestCase(PdfCustomPropertiesExport.Metadata)]
        public void CustomPropertiesExport(PdfCustomPropertiesExport pdfCustomPropertiesExportMode)
        {
            //ExStart
            //ExFor:PdfCustomPropertiesExport
            //ExFor:PdfSaveOptions.CustomPropertiesExport
            //ExSummary:Shows how to export custom properties while converting a document to PDF.
            Document doc = new Document();

            doc.CustomDocumentProperties.Add("Company", "My value");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Set the "CustomPropertiesExport" property to "PdfCustomPropertiesExport.None" to discard
            // custom document properties as we save the document to .PDF. 
            // Set the "CustomPropertiesExport" property to "PdfCustomPropertiesExport.Standard"
            // to preserve custom properties within the output PDF document.
            // Set the "CustomPropertiesExport" property to "PdfCustomPropertiesExport.Metadata"
            // to preserve custom properties in an XMP packet.
            options.CustomPropertiesExport = pdfCustomPropertiesExportMode;

            doc.Save(ArtifactsDir + "PdfSaveOptions.CustomPropertiesExport.pdf", options);
            //ExEnd

            switch (pdfCustomPropertiesExportMode)
            {
                case PdfCustomPropertiesExport.None:
                    if (!IsRunningOnMono())
                    {
                        Assert.Throws<AssertionException>(() => TestUtil.FileContainsString(
                            doc.CustomDocumentProperties[0].Name,
                            ArtifactsDir + "PdfSaveOptions.CustomPropertiesExport.pdf"));
                        Assert.Throws<AssertionException>(() => TestUtil.FileContainsString(
                            "<</Type /Metadata/Subtype /XML/Length 8 0 R/Filter /FlateDecode>>",
                            ArtifactsDir + "PdfSaveOptions.CustomPropertiesExport.pdf"));
                    }

                    break;
                case PdfCustomPropertiesExport.Standard:
                    TestUtil.FileContainsString(
                        "<</Creator(þÿ\0A\0s\0p\0o\0s\0e\0.\0W\0o\0r\0d\0s)/Producer(þÿ\0A\0s\0p\0o\0s\0e\0.\0W\0o\0r\0d\0s\0 \0f\0o\0r\0",
                        ArtifactsDir + "PdfSaveOptions.CustomPropertiesExport.pdf");
                    TestUtil.FileContainsString("/Company (þÿ\0M\0y\0 \0v\0a\0l\0u\0e)>>",
                        ArtifactsDir + "PdfSaveOptions.CustomPropertiesExport.pdf");
                    break;
                case PdfCustomPropertiesExport.Metadata:
                    TestUtil.FileContainsString("<</Type /Metadata/Subtype /XML/Length 8 0 R/Filter /FlateDecode>>",
                        ArtifactsDir + "PdfSaveOptions.CustomPropertiesExport.pdf");
                    break;
            }

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.CustomPropertiesExport.pdf");

            Assert.AreEqual("Aspose.Words", pdfDocument.Info.Creator);
            Assert.True(pdfDocument.Info.Producer.StartsWith("Aspose.Words"));
            
            switch (pdfCustomPropertiesExportMode)
            {
                case PdfCustomPropertiesExport.None:
                    Assert.AreEqual(2, pdfDocument.Info.Count);
                    Assert.AreEqual(0, pdfDocument.Metadata.Count);
                    break;
                case PdfCustomPropertiesExport.Metadata:
                    Assert.AreEqual(2, pdfDocument.Info.Count);
                    Assert.AreEqual(2, pdfDocument.Metadata.Count);

                    Assert.AreEqual("Aspose.Words", pdfDocument.Metadata["xmp:CreatorTool"].ToString());
                    Assert.AreEqual("Company", pdfDocument.Metadata["custprops:Property1"].ToString());
                    break;
                case PdfCustomPropertiesExport.Standard:
                    Assert.AreEqual(3, pdfDocument.Info.Count);
                    Assert.AreEqual(0, pdfDocument.Metadata.Count);

                    Assert.AreEqual("My value", pdfDocument.Info["Company"]);
                    break;
            }
#endif
        }

        [TestCase(DmlEffectsRenderingMode.None)]
        [TestCase(DmlEffectsRenderingMode.Simplified)]
        [TestCase(DmlEffectsRenderingMode.Fine)]
        public void DrawingMLEffects(DmlEffectsRenderingMode effectsRenderingMode)
        {
            //ExStart
            //ExFor:DmlRenderingMode
            //ExFor:DmlEffectsRenderingMode
            //ExFor:PdfSaveOptions.DmlEffectsRenderingMode
            //ExFor:SaveOptions.DmlEffectsRenderingMode
            //ExFor:SaveOptions.DmlRenderingMode
            //ExSummary:Shows how to configure the rendering quality of DrawingML effects in a document as we save it to PDF.
            Document doc = new Document(MyDir + "DrawingML shape effects.docx");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Set the "DmlEffectsRenderingMode" property to "DmlEffectsRenderingMode.None" to discard all DrawingML effects.
            // Set the "DmlEffectsRenderingMode" property to "DmlEffectsRenderingMode.Simplified"
            // to render a simplified version of DrawingML effects.
            // Set the "DmlEffectsRenderingMode" property to "DmlEffectsRenderingMode.Fine" to
            // render DrawingML effects with more accuracy and also with more processing cost.
            options.DmlEffectsRenderingMode = effectsRenderingMode;

            Assert.AreEqual(DmlRenderingMode.DrawingML, options.DmlRenderingMode);

            doc.Save(ArtifactsDir + "PdfSaveOptions.DrawingMLEffects.pdf", options);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument =
                new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.DrawingMLEffects.pdf");

            ImagePlacementAbsorber imagePlacementAbsorber = new ImagePlacementAbsorber();
            imagePlacementAbsorber.Visit(pdfDocument.Pages[1]);

            TableAbsorber tableAbsorber = new TableAbsorber();
            tableAbsorber.Visit(pdfDocument.Pages[1]);

            switch (effectsRenderingMode)
            {
                case DmlEffectsRenderingMode.None:
                case DmlEffectsRenderingMode.Simplified:
                    TestUtil.FileContainsString("4 0 obj\r\n" +
                                                "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                        ArtifactsDir + "PdfSaveOptions.DrawingMLEffects.pdf");
                    Assert.AreEqual(0, imagePlacementAbsorber.ImagePlacements.Count);
                    Assert.AreEqual(28, tableAbsorber.TableList.Count);
                    break;
                case DmlEffectsRenderingMode.Fine:
                    TestUtil.FileContainsString(
                        "4 0 obj\r\n<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R>>/XObject<</X1 9 0 R/X2 10 0 R/X3 11 0 R/X4 12 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                        ArtifactsDir + "PdfSaveOptions.DrawingMLEffects.pdf");
                    Assert.AreEqual(21, imagePlacementAbsorber.ImagePlacements.Count);
                    Assert.AreEqual(4, tableAbsorber.TableList.Count);
                    break;
            }
#endif
        }

        [TestCase(DmlRenderingMode.Fallback)]
        [TestCase(DmlRenderingMode.DrawingML)]
        public void DrawingMLFallback(DmlRenderingMode dmlRenderingMode)
        {
            //ExStart
            //ExFor:DmlRenderingMode
            //ExFor:SaveOptions.DmlRenderingMode
            //ExSummary:Shows how to render fallback shapes when saving to PDF.
            Document doc = new Document(MyDir + "DrawingML shape fallbacks.docx");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Set the "DmlRenderingMode" property to "DmlRenderingMode.Fallback"
            // to substitute DML shapes with their fallback shapes.
            // Set the "DmlRenderingMode" property to "DmlRenderingMode.DrawingML"
            // to render the DML shapes themselves.
            options.DmlRenderingMode = dmlRenderingMode;

            doc.Save(ArtifactsDir + "PdfSaveOptions.DrawingMLFallback.pdf", options);
            //ExEnd

            switch (dmlRenderingMode)
            {
                case DmlRenderingMode.DrawingML:
                    TestUtil.FileContainsString(
                        "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R/FAAABA 10 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                        ArtifactsDir + "PdfSaveOptions.DrawingMLFallback.pdf");
                    break;
                case DmlRenderingMode.Fallback:
                    TestUtil.FileContainsString(
                        "4 0 obj\r\n<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R/FAAABC 12 0 R>>/ExtGState<</GS1 9 0 R/GS2 10 0 R>>>>/Group ",
                        ArtifactsDir + "PdfSaveOptions.DrawingMLFallback.pdf");
                    break;
            }

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument =
                new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.DrawingMLFallback.pdf");

            ImagePlacementAbsorber imagePlacementAbsorber = new ImagePlacementAbsorber();
            imagePlacementAbsorber.Visit(pdfDocument.Pages[1]);

            TableAbsorber tableAbsorber = new TableAbsorber();
            tableAbsorber.Visit(pdfDocument.Pages[1]);

            switch (dmlRenderingMode)
            {
                case DmlRenderingMode.DrawingML:
                    Assert.AreEqual(6, tableAbsorber.TableList.Count);
                    break;
                case DmlRenderingMode.Fallback:
                    Assert.AreEqual(15, tableAbsorber.TableList.Count);
                    break;
            }
#endif
        }

        [TestCase(false)]
        [TestCase(true)]
        public void ExportDocumentStructure(bool exportDocumentStructure)
        {
            //ExStart
            //ExFor:PdfSaveOptions.ExportDocumentStructure
            //ExSummary:Shows how to preserve document structure elements, which can assist in programmatically interpreting our document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ParagraphFormat.Style = doc.Styles["Heading 1"];
            builder.Writeln("Hello world!");
            builder.ParagraphFormat.Style = doc.Styles["Normal"];
            builder.Write(
                "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Set the "ExportDocumentStructure" property to "true" to make the document structure, such tags, available via the
            // "Content" navigation pane of Adobe Acrobat at the cost of increased file size.
            // Set the "ExportDocumentStructure" property to "false" to not export the document structure.
            options.ExportDocumentStructure = exportDocumentStructure;

            // Suppose we export document structure while saving this document. In that case,
            // we can open it using Adobe Acrobat and find tags for elements such as the heading
            // and the next paragraph via "View" -> "Show/Hide" -> "Navigation panes" -> "Tags".
            doc.Save(ArtifactsDir + "PdfSaveOptions.ExportDocumentStructure.pdf", options);
            //ExEnd

            if (exportDocumentStructure)
            {
                TestUtil.FileContainsString("4 0 obj\r\n" +
                                            "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R/FAAABB 11 0 R>>/ExtGState<</GS1 9 0 R/GS2 13 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>/StructParents 0/Tabs /S>>",
                    ArtifactsDir + "PdfSaveOptions.ExportDocumentStructure.pdf");
            }
            else
            {
                TestUtil.FileContainsString("4 0 obj\r\n" +
                                            "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 612 792]/Resources<</Font<</FAAAAH 7 0 R/FAAABA 10 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                    ArtifactsDir + "PdfSaveOptions.ExportDocumentStructure.pdf");
            }
        }

#if NET462 || JAVA
        [TestCase(false, Category = "SkipMono")]
        [TestCase(true, Category = "SkipMono")]
        public void PreblendImages(bool preblendImages)
        {
            //ExStart
            //ExFor:PdfSaveOptions.PreblendImages
            //ExSummary:Shows how to preblend images with transparent backgrounds while saving a document to PDF.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Image img = Image.FromFile(ImageDir + "Transparent background logo.png");
            builder.InsertImage(img);

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Set the "PreblendImages" property to "true" to preblend transparent images
            // with a background, which may reduce artifacts.
            // Set the "PreblendImages" property to "false" to render transparent images normally.
            options.PreblendImages = preblendImages;

            doc.Save(ArtifactsDir + "PdfSaveOptions.PreblendImages.pdf", options);
            //ExEnd

            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.PreblendImages.pdf");
            XImage image = pdfDocument.Pages[1].Resources.Images[1];

            using (MemoryStream stream = new MemoryStream())
            {
                image.Save(stream);

                if (preblendImages)
                {
                    TestUtil.FileContainsString("9 0 obj\r\n20849 ", ArtifactsDir + "PdfSaveOptions.PreblendImages.pdf");
                    Assert.AreEqual(17898, stream.Length);
                }
                else
                {
                    TestUtil.FileContainsString("9 0 obj\r\n19289 ", ArtifactsDir + "PdfSaveOptions.PreblendImages.pdf");
                    Assert.AreEqual(19216, stream.Length);
                }
            }
        }

        [TestCase(false)]
        [TestCase(true)]
        public void InterpolateImages(bool interpolateImages)
        {
            //ExStart
            //ExFor:PdfSaveOptions.InterpolateImages
            //ExSummary:Shows how to perform interpolation on images while saving a document to PDF.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Image img = Image.FromFile(ImageDir + "Transparent background logo.png");
            builder.InsertImage(img);

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Set the "InterpolateImages" property to "true" to get the reader that opens this document to interpolate images.
            // Their resolution should be lower than that of the device that is displaying the document.
            // Set the "InterpolateImages" property to "false" to make it so that the reader does not apply any interpolation.
            saveOptions.InterpolateImages = interpolateImages;

            // When we open this document with a reader such as Adobe Acrobat, we will need to zoom in on the image
            // to see the interpolation effect if we saved the document with it enabled.
            doc.Save(ArtifactsDir + "PdfSaveOptions.InterpolateImages.pdf", saveOptions);
            //ExEnd

            if (interpolateImages)
            {
                TestUtil.FileContainsString("6 0 obj\r\n" +
                                            "<</Type /XObject/Subtype /Image/Width 400/Height 400/ColorSpace /DeviceRGB/BitsPerComponent 8/SMask 8 0 R/Interpolate true/Length 9 0 R/Filter /FlateDecode>>",
                    ArtifactsDir + "PdfSaveOptions.InterpolateImages.pdf");
            }
            else
            {
                TestUtil.FileContainsString("6 0 obj\r\n" +
                                            "<</Type /XObject/Subtype /Image/Width 400/Height 400/ColorSpace /DeviceRGB/BitsPerComponent 8/SMask 8 0 R/Length 9 0 R/Filter /FlateDecode>>",
                    ArtifactsDir + "PdfSaveOptions.InterpolateImages.pdf");
            }
        }

        [Test, Category("SkipMono")]
        public void Dml3DEffectsRenderingModeTest()
        {
            Document doc = new Document(MyDir + "DrawingML shape 3D effects.docx");
            
            RenderCallback warningCallback = new RenderCallback();
            doc.WarningCallback = warningCallback;
            
            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.Dml3DEffectsRenderingMode = Dml3DEffectsRenderingMode.Advanced;
            
            doc.Save(ArtifactsDir + "PdfSaveOptions.Dml3DEffectsRenderingModeTest.pdf", saveOptions);

            Assert.AreEqual(43, warningCallback.Count);
        }

        public class RenderCallback : IWarningCallback
        {
            public void Warning(WarningInfo info)
            {
                Console.WriteLine($"{info.WarningType}: {info.Description}.");
                mWarnings.Add(info);
            }

            public WarningInfo this[int i] => mWarnings[i];

            /// <summary>
            /// Clears warning collection.
            /// </summary>
            public void Clear()
            {
                mWarnings.Clear();
            }

            public int Count => mWarnings.Count;

            /// <summary>
            /// Returns true if a warning with the specified properties has been generated.
            /// </summary>
            public bool Contains(WarningSource source, WarningType type, string description)
            {
                return mWarnings.Any(warning =>
                    warning.Source == source && warning.WarningType == type && warning.Description == description);
            }

            private readonly List<WarningInfo> mWarnings = new List<WarningInfo>();
        }

#elif NETCOREAPP2_1
        [TestCase(false)]
        [TestCase(true)]
        public void PreblendImagesNetStandard2(bool preblendImages)
        {
            //ExStart
            //ExFor:PdfSaveOptions.PreblendImages
            //ExSummary:Shows how to preblend images with transparent backgrounds (.NetStandard 2.0).
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            using (Image image = Image.Decode(ImageDir + "Transparent background logo.png"))
                builder.InsertImage(image);

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Set the "PreblendImages" property to "true" to preblend transparent images
            // with a background, which may reduce artifacts.
            // Set the "PreblendImages" property to "false" to render transparent images normally.
            options.PreblendImages = preblendImages;

            doc.Save(ArtifactsDir + "PdfSaveOptions.PreblendImagesNetStandard2.pdf", options);
            //ExEnd

            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.PreblendImagesNetStandard2.pdf");
            XImage xImage = pdfDocument.Pages[1].Resources.Images[1];

            using (MemoryStream stream = new MemoryStream())
            {
                xImage.Save(stream);

                if (preblendImages)
                {
                    TestUtil.FileContainsString("9 0 obj\r\n20849 ", ArtifactsDir + "PdfSaveOptions.PreblendImagesNetStandard2.pdf");
                    Assert.AreEqual(17898, stream.Length);
                }
                else
                {
                    TestUtil.FileContainsString("9 0 obj\r\n20266 ", ArtifactsDir + "PdfSaveOptions.PreblendImagesNetStandard2.pdf");
                    Assert.AreEqual(19135, stream.Length);
                }
            }
        }

        [TestCase(false)]
        [TestCase(true)]
        public void InterpolateImagesNetStandard2(bool interpolateImages)
        {
            //ExStart
            //ExFor:PdfSaveOptions.InterpolateImages
            //ExSummary:Shows how to improve the quality of an image in the rendered documents (.NetStandard 2.0).
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            using (Image image = Image.Decode(ImageDir + "Transparent background logo.png"))
                builder.InsertImage(image);

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Set the "InterpolateImages" property to "true" to get the reader that opens this document to interpolate images.
            // Their resolution should be lower than that of the device that is displaying the document.
            // Set the "InterpolateImages" property to "false" to make it so that the reader does not apply any interpolation.
            saveOptions.InterpolateImages = interpolateImages;

            // When we open this document with a reader such as Adobe Acrobat, we will need to zoom in on the image
            // to see the interpolation effect if we saved the document with it enabled.
            doc.Save(ArtifactsDir + "PdfSaveOptions.InterpolateImagesNetStandard2.pdf", saveOptions);
            //ExEnd

            if (interpolateImages)
            {
                TestUtil.FileContainsString("6 0 obj\r\n" +
                                            "<</Type /XObject/Subtype /Image/Width 400/Height 400/ColorSpace /DeviceRGB/BitsPerComponent 8/SMask 8 0 R/Interpolate true/Length 9 0 R/Filter /FlateDecode>>",
                    ArtifactsDir + "PdfSaveOptions.InterpolateImagesNetStandard2.pdf");
            }
            else
            {
                TestUtil.FileContainsString("6 0 obj\r\n" +
                                            "<</Type /XObject/Subtype /Image/Width 400/Height 400/ColorSpace /DeviceRGB/BitsPerComponent 8/SMask 8 0 R/Length 9 0 R/Filter /FlateDecode>>",
                    ArtifactsDir + "PdfSaveOptions.InterpolateImagesNetStandard2.pdf");
            }
        }
#endif

        [Test]
        public void PdfDigitalSignature()
        {
            //ExStart
            //ExFor:PdfDigitalSignatureDetails
            //ExFor:PdfDigitalSignatureDetails.#ctor
            //ExFor:PdfDigitalSignatureDetails.#ctor(CertificateHolder, String, String, DateTime)
            //ExFor:PdfDigitalSignatureDetails.HashAlgorithm
            //ExFor:PdfDigitalSignatureDetails.Location
            //ExFor:PdfDigitalSignatureDetails.Reason
            //ExFor:PdfDigitalSignatureDetails.SignatureDate
            //ExFor:PdfDigitalSignatureHashAlgorithm
            //ExFor:PdfSaveOptions.DigitalSignatureDetails
            //ExSummary:Shows how to sign a generated PDF document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Contents of signed PDF.");

            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Configure the "DigitalSignatureDetails" object of the "SaveOptions" object to
            // digitally sign the document as we render it with the "Save" method.
            DateTime signingTime = DateTime.Now;
            options.DigitalSignatureDetails =
                new PdfDigitalSignatureDetails(certificateHolder, "Test Signing", "My Office", signingTime);
            options.DigitalSignatureDetails.HashAlgorithm = PdfDigitalSignatureHashAlgorithm.Sha256;

            Assert.AreEqual("Test Signing", options.DigitalSignatureDetails.Reason);
            Assert.AreEqual("My Office", options.DigitalSignatureDetails.Location);
            Assert.AreEqual(signingTime.ToUniversalTime(), options.DigitalSignatureDetails.SignatureDate);

            doc.Save(ArtifactsDir + "PdfSaveOptions.PdfDigitalSignature.pdf", options);
            //ExEnd

            TestUtil.FileContainsString("6 0 obj\r\n" +
                                        "<</Type /Annot/Subtype /Widget/FT /Sig/DR <<>>/F 132/Rect [0 0 0 0]/V 7 0 R/P 4 0 R/T(þÿ\0A\0s\0p\0o\0s\0e\0D\0i\0g\0i\0t\0a\0l\0S\0i\0g\0n\0a\0t\0u\0r\0e)/AP <</N 8 0 R>>>>",
                ArtifactsDir + "PdfSaveOptions.PdfDigitalSignature.pdf");

            Assert.False(FileFormatUtil.DetectFileFormat(ArtifactsDir + "PdfSaveOptions.PdfDigitalSignature.pdf")
                .HasDigitalSignature);

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.PdfDigitalSignature.pdf");

            Assert.False(pdfDocument.Form.SignaturesExist);

            SignatureField signatureField = (SignatureField)pdfDocument.Form[1];

            Assert.AreEqual("AsposeDigitalSignature", signatureField.FullName);
            Assert.AreEqual("AsposeDigitalSignature", signatureField.PartialName);
            Assert.AreEqual(typeof(Aspose.Pdf.Forms.PKCS7), signatureField.Signature.GetType());
            Assert.AreEqual(DateTime.Today, signatureField.Signature.Date.Date);
            Assert.AreEqual("þÿ\0M\0o\0r\0z\0a\0l\0.\0M\0e", signatureField.Signature.Authority);
            Assert.AreEqual("þÿ\0M\0y\0 \0O\0f\0f\0i\0c\0e", signatureField.Signature.Location);
            Assert.AreEqual("þÿ\0T\0e\0s\0t\0 \0S\0i\0g\0n\0i\0n\0g", signatureField.Signature.Reason);
#endif
        }
        
        [Test]
        public void PdfDigitalSignatureTimestamp()
        {
            //ExStart
            //ExFor:PdfDigitalSignatureDetails.TimestampSettings
            //ExFor:PdfDigitalSignatureTimestampSettings
            //ExFor:PdfDigitalSignatureTimestampSettings.#ctor
            //ExFor:PdfDigitalSignatureTimestampSettings.#ctor(String,String,String)
            //ExFor:PdfDigitalSignatureTimestampSettings.#ctor(String,String,String,TimeSpan)
            //ExFor:PdfDigitalSignatureTimestampSettings.Password
            //ExFor:PdfDigitalSignatureTimestampSettings.ServerUrl
            //ExFor:PdfDigitalSignatureTimestampSettings.Timeout
            //ExFor:PdfDigitalSignatureTimestampSettings.UserName
            //ExSummary:Shows how to sign a saved PDF document digitally and timestamp it.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Signed PDF contents.");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Create a digital signature and assign it to our SaveOptions object to sign the document when we save it to PDF. 
            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");
            options.DigitalSignatureDetails = new PdfDigitalSignatureDetails(certificateHolder, "Test Signing", "Aspose Office", DateTime.Now);

            // Create a timestamp authority-verified timestamp.
            options.DigitalSignatureDetails.TimestampSettings =
                new PdfDigitalSignatureTimestampSettings("https://freetsa.org/tsr", "JohnDoe", "MyPassword");

            // The default lifespan of the timestamp is 100 seconds.
            Assert.AreEqual(100.0d, options.DigitalSignatureDetails.TimestampSettings.Timeout.TotalSeconds);

            // We can set our timeout period via the constructor.
            options.DigitalSignatureDetails.TimestampSettings =
                new PdfDigitalSignatureTimestampSettings("https://freetsa.org/tsr", "JohnDoe", "MyPassword", TimeSpan.FromMinutes(30));

            Assert.AreEqual(1800.0d, options.DigitalSignatureDetails.TimestampSettings.Timeout.TotalSeconds);
            Assert.AreEqual("https://freetsa.org/tsr", options.DigitalSignatureDetails.TimestampSettings.ServerUrl);
            Assert.AreEqual("JohnDoe", options.DigitalSignatureDetails.TimestampSettings.UserName);
            Assert.AreEqual("MyPassword", options.DigitalSignatureDetails.TimestampSettings.Password);

            // The "Save" method will apply our signature to the output document at this time.
            doc.Save(ArtifactsDir + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf", options);
            //ExEnd

            Assert.False(FileFormatUtil.DetectFileFormat(ArtifactsDir + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf").HasDigitalSignature);
            TestUtil.FileContainsString("6 0 obj\r\n" +
                                        "<</Type /Annot/Subtype /Widget/FT /Sig/DR <<>>/F 132/Rect [0 0 0 0]/V 7 0 R/P 4 0 R/T(þÿ\0A\0s\0p\0o\0s\0e\0D\0i\0g\0i\0t\0a\0l\0S\0i\0g\0n\0a\0t\0u\0r\0e)/AP <</N 8 0 R>>>>", 
            ArtifactsDir + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf");

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.PdfDigitalSignatureTimestamp.pdf");

            Assert.False(pdfDocument.Form.SignaturesExist);

            SignatureField signatureField = (SignatureField)pdfDocument.Form[1];

            Assert.AreEqual("AsposeDigitalSignature", signatureField.FullName);
            Assert.AreEqual("AsposeDigitalSignature", signatureField.PartialName);
            Assert.AreEqual(typeof(Aspose.Pdf.Forms.PKCS7), signatureField.Signature.GetType());
            Assert.AreEqual(new DateTime(1, 1, 1, 0, 0, 0), signatureField.Signature.Date);
            Assert.AreEqual("þÿ\0M\0o\0r\0z\0a\0l\0.\0M\0e", signatureField.Signature.Authority);
            Assert.AreEqual("þÿ\0A\0s\0p\0o\0s\0e\0 \0O\0f\0f\0i\0c\0e", signatureField.Signature.Location);
            Assert.AreEqual("þÿ\0T\0e\0s\0t\0 \0S\0i\0g\0n\0i\0n\0g", signatureField.Signature.Reason);
            Assert.Null(signatureField.Signature.TimestampSettings);
#endif
        }

        [TestCase(EmfPlusDualRenderingMode.Emf)]
        [TestCase(EmfPlusDualRenderingMode.EmfPlus)]
        [TestCase(EmfPlusDualRenderingMode.EmfPlusWithFallback)]
        public void RenderMetafile(EmfPlusDualRenderingMode renderingMode)
        {
            //ExStart
            //ExFor:EmfPlusDualRenderingMode
            //ExFor:MetafileRenderingOptions.EmfPlusDualRenderingMode
            //ExFor:MetafileRenderingOptions.UseEmfEmbeddedToWmf
            //ExSummary:Shows how to configure Enhanced Windows Metafile-related rendering options when saving to PDF.
            Document doc = new Document(MyDir + "EMF.docx");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Set the "EmfPlusDualRenderingMode" property to "EmfPlusDualRenderingMode.Emf"
            // to only render the EMF part of an EMF+ dual metafile.
            // Set the "EmfPlusDualRenderingMode" property to "EmfPlusDualRenderingMode.EmfPlus" to
            // to render the EMF+ part of an EMF+ dual metafile.
            // Set the "EmfPlusDualRenderingMode" property to "EmfPlusDualRenderingMode.EmfPlusWithFallback"
            // to render the EMF+ part of an EMF+ dual metafile if all of the EMF+ records are supported.
            // Otherwise, Aspose.Words will render the EMF part.
            saveOptions.MetafileRenderingOptions.EmfPlusDualRenderingMode = renderingMode;

            // Set the "UseEmfEmbeddedToWmf" property to "true" to render embedded EMF data
            // for metafiles that we can render as vector graphics.
            saveOptions.MetafileRenderingOptions.UseEmfEmbeddedToWmf = true;

            doc.Save(ArtifactsDir + "PdfSaveOptions.RenderMetafile.pdf", saveOptions);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument =
                new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.RenderMetafile.pdf");

            switch (renderingMode)
            {
                case EmfPlusDualRenderingMode.Emf:
                case EmfPlusDualRenderingMode.EmfPlusWithFallback:
                    Assert.AreEqual(0, pdfDocument.Pages[1].Resources.Images.Count);
                    TestUtil.FileContainsString("4 0 obj\r\n" +
                                                "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 595.29998779 841.90002441]/Resources<</Font<</FAAAAH 7 0 R/FAAABA 10 0 R/FAAABD 13 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                        ArtifactsDir + "PdfSaveOptions.RenderMetafile.pdf");
                    break;
                case EmfPlusDualRenderingMode.EmfPlus:
                    Assert.AreEqual(1, pdfDocument.Pages[1].Resources.Images.Count);
                    TestUtil.FileContainsString("4 0 obj\r\n" +
                                                "<</Type /Page/Parent 3 0 R/Contents 5 0 R/MediaBox [0 0 595.29998779 841.90002441]/Resources<</Font<</FAAAAH 7 0 R/FAAABB 11 0 R/FAAABE 14 0 R>>/XObject<</X1 9 0 R>>>>/Group <</Type/Group/S/Transparency/CS/DeviceRGB>>>>",
                        ArtifactsDir + "PdfSaveOptions.RenderMetafile.pdf");
                    break;
            }
#endif
        }

        [Test]
        public void EncryptionPermissions()
        {
            //ExStart
            //ExFor:PdfEncryptionDetails.#ctor
            //ExFor:PdfSaveOptions.EncryptionDetails
            //ExFor:PdfEncryptionDetails.Permissions
            //ExFor:PdfEncryptionDetails.EncryptionAlgorithm
            //ExFor:PdfEncryptionDetails.OwnerPassword
            //ExFor:PdfEncryptionDetails.UserPassword
            //ExFor:PdfEncryptionAlgorithm
            //ExFor:PdfPermissions
            //ExFor:PdfEncryptionDetails
            //ExSummary:Shows how to set permissions on a saved PDF document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");

            PdfEncryptionDetails encryptionDetails =
                new PdfEncryptionDetails("password", string.Empty, PdfEncryptionAlgorithm.RC4_128);

            // Start by disallowing all permissions.
            encryptionDetails.Permissions = PdfPermissions.DisallowAll;

            // Extend permissions to allow the editing of annotations.
            encryptionDetails.Permissions = PdfPermissions.ModifyAnnotations | PdfPermissions.DocumentAssembly;

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Enable encryption via the "EncryptionDetails" property.
            saveOptions.EncryptionDetails = encryptionDetails;

            // When we open this document, we will need to provide the password before accessing its contents.
            doc.Save(ArtifactsDir + "PdfSaveOptions.EncryptionPermissions.pdf", saveOptions);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument;

            Assert.Throws<InvalidPasswordException>(() => 
                pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.EncryptionPermissions.pdf"));

            pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.EncryptionPermissions.pdf", "password");
            TextFragmentAbsorber textAbsorber = new TextFragmentAbsorber();

            pdfDocument.Pages[1].Accept(textAbsorber);
            
            Assert.AreEqual("Hello world!", textAbsorber.Text);
#endif
        }

        [TestCase(NumeralFormat.ArabicIndic)]
        [TestCase(NumeralFormat.Context)]
        [TestCase(NumeralFormat.EasternArabicIndic)]
        [TestCase(NumeralFormat.European)]
        [TestCase(NumeralFormat.System)]
        public void SetNumeralFormat(NumeralFormat numeralFormat)
        {
            //ExStart
            //ExFor:FixedPageSaveOptions.NumeralFormat
            //ExFor:NumeralFormat
            //ExSummary:Shows how to set the numeral format used when saving to PDF.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.LocaleId = new CultureInfo("ar-AR").LCID;
            builder.Writeln("1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 50, 100");

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Set the "NumeralFormat" property to "NumeralFormat.ArabicIndic" to
            // use glyphs from the U+0660 to U+0669 range as numbers.
            // Set the "NumeralFormat" property to "NumeralFormat.Context" to
            // look up the locale to determine what number of glyphs to use.
            // Set the "NumeralFormat" property to "NumeralFormat.EasternArabicIndic" to
            // use glyphs from the U+06F0 to U+06F9 range as numbers.
            // Set the "NumeralFormat" property to "NumeralFormat.European" to use european numerals.
            // Set the "NumeralFormat" property to "NumeralFormat.System" to determine the symbol set from regional settings.
            options.NumeralFormat = numeralFormat;

            doc.Save(ArtifactsDir + "PdfSaveOptions.SetNumeralFormat.pdf", options);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.SetNumeralFormat.pdf");
            TextFragmentAbsorber textAbsorber = new TextFragmentAbsorber();

            pdfDocument.Pages[1].Accept(textAbsorber);

            switch (numeralFormat)
            {
                case NumeralFormat.European:
                    Assert.AreEqual("1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 50, 100", textAbsorber.Text);
                    break;
                case NumeralFormat.ArabicIndic:
                    Assert.AreEqual(", ٢, ٣, ٤, ٥, ٦, ٧, ٨, ٩, ١٠, ٥٠, ١١٠٠", textAbsorber.Text);
                    break;
                case NumeralFormat.EasternArabicIndic:
                    Assert.AreEqual("۱۰۰ ,۵۰ ,۱۰ ,۹ ,۸ ,۷ ,۶ ,۵ ,۴ ,۳ ,۲ ,۱", textAbsorber.Text);
                    break;
            }
#endif
        }

        [Test]
        public void ExportPageSet()
        {
            //ExStart
            //ExFor:FixedPageSaveOptions.PageSet
            //ExSummary:Shows how to export Odd pages from the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            for (int i = 0; i < 5; i++)
            {
                builder.Writeln($"Page {i + 1} ({(i % 2 == 0 ? "odd" : "even")})");
                if (i < 4)
                    builder.InsertBreak(BreakType.PageBreak);
            }

            // Create a "PdfSaveOptions" object that we can pass to the document's "Save" method
            // to modify how that method converts the document to .PDF.
            PdfSaveOptions options = new PdfSaveOptions();

            // Below are three PageSet properties that we can use to filter out a set of pages from
            // our document to save in an output PDF document based on the parity of their page numbers.
            // 1 -  Save only the even-numbered pages:
            options.PageSet = PageSet.Even;

            doc.Save(ArtifactsDir + "PdfSaveOptions.ExportPageSet.Even.pdf", options);

            // 2 -  Save only the odd-numbered pages:
            options.PageSet = PageSet.Odd;

            doc.Save(ArtifactsDir + "PdfSaveOptions.ExportPageSet.Odd.pdf", options);

            // 3 -  Save every page:
            options.PageSet = PageSet.All;

            doc.Save(ArtifactsDir + "PdfSaveOptions.ExportPageSet.All.pdf", options);
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.ExportPageSet.Even.pdf");
            TextAbsorber textAbsorber = new TextAbsorber();
            pdfDocument.Pages.Accept(textAbsorber);

            Assert.AreEqual("Page 2 (even)\r\n" +
                            "Page 4 (even)", textAbsorber.Text);

            pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.ExportPageSet.Odd.pdf");
            textAbsorber = new TextAbsorber();
            pdfDocument.Pages.Accept(textAbsorber);

            Assert.AreEqual("Page 1 (odd)\r\n" +
                            "Page 3 (odd)\r\n" +
                            "Page 5 (odd)", textAbsorber.Text);

            pdfDocument = new Aspose.Pdf.Document(ArtifactsDir + "PdfSaveOptions.ExportPageSet.All.pdf");
            textAbsorber = new TextAbsorber();
            pdfDocument.Pages.Accept(textAbsorber);

            Assert.AreEqual("Page 1 (odd)\r\n" +
                            "Page 2 (even)\r\n" +
                            "Page 3 (odd)\r\n" +
                            "Page 4 (even)\r\n" +
                            "Page 5 (odd)", textAbsorber.Text);
#endif
        }
    }
}