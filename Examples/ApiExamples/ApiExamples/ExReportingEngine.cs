﻿// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using ApiExamples.TestData;
using ApiExamples.TestData.TestBuilders;
using ApiExamples.TestData.TestClasses;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;
using Aspose.Words.Reporting;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExReportingEngine : ApiExampleBase
    {
        private readonly string mImage = ImageDir + "Logo.jpg";
        private readonly string mDocument = MyDir + "Reporting engine template - Data table.docx";

        [Test]
        public void SimpleCase()
        {
            Document doc = DocumentHelper.CreateSimpleDocument("<<[s.Name]>> says: <<[s.Message]>>");

            MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "Hello World");
            BuildReport(doc, sender, "s", ReportBuildOptions.InlineErrorMessages);

            doc = DocumentHelper.SaveOpen(doc);

            Assert.That(doc.GetText(), Is.EqualTo("LINQ Reporting Engine says: Hello World\f"));
        }

        [Test]
        public void StringFormat()
        {
            Document doc = DocumentHelper.CreateSimpleDocument(
                "<<[s.Name]:lower>> says: <<[s.Message]:upper>>, <<[s.Message]:caps>>, <<[s.Message]:firstCap>>");

            MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "hello world");
            BuildReport(doc, sender, "s");

            doc = DocumentHelper.SaveOpen(doc);

            Assert.That(doc.GetText(), Is.EqualTo("linq reporting engine says: HELLO WORLD, Hello World, Hello world\f"));
        }

        [Test]
        public void NumberFormat()
        {
            Document doc = DocumentHelper.CreateSimpleDocument(
                "<<[s.Value1]:alphabetic>> : <<[s.Value2]:roman:lower>>, <<[s.Value3]:ordinal>>, <<[s.Value1]:ordinalText:upper>>" +
                ", <<[s.Value2]:cardinal>>, <<[s.Value3]:hex>>, <<[s.Value3]:arabicDash>>");

            NumericTestClass sender = new NumericTestBuilder()
                .WithValuesAndDate(1, 2.2, 200, null, DateTime.Parse("10.09.2016 10:00:00")).Build();
            BuildReport(doc, sender, "s");

            doc = DocumentHelper.SaveOpen(doc);

            Assert.That(doc.GetText(), Is.EqualTo("A : ii, 200th, FIRST, Two, C8, - 200 -\f"));
        }

        [Test]
        public void TestDataTable()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Data table.docx");

            BuildReport(doc, Common.GetContracts(), "Contracts");

            doc.Save(ArtifactsDir + "ReportingEngine.TestDataTable.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.TestDataTable.docx", GoldsDir + "ReportingEngine.TestDataTable Gold.docx"), Is.True);
        }

        [Test]
        public void Total()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Total.docx");

            BuildReport(doc, Common.GetContracts(), "Contracts");

            doc.Save(ArtifactsDir + "ReportingEngine.Total.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.Total.docx", GoldsDir + "ReportingEngine.Total Gold.docx"), Is.True);
        }

        [Test]
        public void TestNestedDataTable()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Nested data table.docx");

            BuildReport(doc, Common.GetManagers(), "Managers");

            doc.Save(ArtifactsDir + "ReportingEngine.TestNestedDataTable.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.TestNestedDataTable.docx", GoldsDir + "ReportingEngine.TestNestedDataTable Gold.docx"), Is.True);
        }

        [Test]
        public void RestartingListNumberingDynamically()
        {
            Document template = new Document(MyDir + "Reporting engine template - List numbering.docx");

            BuildReport(template, Common.GetManagers(), "Managers", ReportBuildOptions.RemoveEmptyParagraphs);

            template.Save(ArtifactsDir + "ReportingEngine.RestartingListNumberingDynamically.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.RestartingListNumberingDynamically.docx", GoldsDir + "ReportingEngine.RestartingListNumberingDynamically Gold.docx"), Is.True);
        }

        [Test]
        public void RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically()
        {
            Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.Document] -build>>");
            
            DocumentTestClass doc = new DocumentTestBuilder()
                .WithDocument(new Document(MyDir + "Reporting engine template - List numbering.docx")).Build();

            BuildReport(template, new object[] {doc, Common.GetManagers()} , new[] {"src", "Managers"}, ReportBuildOptions.RemoveEmptyParagraphs);

            template.Save(ArtifactsDir + "ReportingEngine.RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically.docx", GoldsDir + "ReportingEngine.RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically Gold.docx"), Is.True);
        }

        [Test]
        public void RestartingListNumberingDynamicallyWhileMultipleInsertionsDocumentDynamically()
        {
            Document mainTemplate = DocumentHelper.CreateSimpleDocument("<<doc [src] -build>>");
            Document template1 = DocumentHelper.CreateSimpleDocument("<<doc [src1] -build>>");
            Document template2 = DocumentHelper.CreateSimpleDocument("<<doc [src2.Document] -build>>");
            
            DocumentTestClass doc = new DocumentTestBuilder()
                .WithDocument(new Document(MyDir + "Reporting engine template - List numbering.docx")).Build();

            BuildReport(mainTemplate, new object[] {template1, template2, doc, Common.GetManagers()} , new[] {"src", "src1", "src2", "Managers"}, ReportBuildOptions.RemoveEmptyParagraphs);

            mainTemplate.Save(ArtifactsDir + "ReportingEngine.RestartingListNumberingDynamicallyWhileMultipleInsertionsDocumentDynamically.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.RestartingListNumberingDynamicallyWhileMultipleInsertionsDocumentDynamically.docx", GoldsDir + "ReportingEngine.RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically Gold.docx"), Is.True);
        }

        [Test]
        public void ChartTest()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Chart.docx");

            BuildReport(doc, Common.GetManagers(), "managers");

            doc.Save(ArtifactsDir + "ReportingEngine.TestChart.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.TestChart.docx", GoldsDir + "ReportingEngine.TestChart Gold.docx"), Is.True);
        }

        [Test]
        public void BubbleChartTest()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Bubble chart.docx");

            BuildReport(doc, Common.GetManagers(), "managers");

            doc.Save(ArtifactsDir + "ReportingEngine.TestBubbleChart.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.TestBubbleChart.docx", GoldsDir + "ReportingEngine.TestBubbleChart Gold.docx"), Is.True);
        }

        [Test]
        public void SetChartSeriesColorsDynamically()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Chart series color.docx");

            BuildReport(doc, Common.GetManagers(), "managers");

            doc.Save(ArtifactsDir + "ReportingEngine.SetChartSeriesColorDynamically.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.SetChartSeriesColorDynamically.docx", GoldsDir + "ReportingEngine.SetChartSeriesColorDynamically Gold.docx"), Is.True);
        }

        [Test]
        public void SetPointColorsDynamically()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Point color.docx");

            List<ColorItemTestClass> colors = new List<ColorItemTestClass>
            {
                new ColorItemTestBuilder().WithColorCodeAndValues("Black", Color.Black.ToArgb(), 1.0, 2.5, 3.5).Build(),
                new ColorItemTestBuilder().WithColorCodeAndValues("Red", Color.Red.ToArgb(), 2.0, 4.0, 2.5).Build(),
                new ColorItemTestBuilder().WithColorCodeAndValues("Green", Color.Green.ToArgb(), 0.5, 1.5, 2.5).Build(),
                new ColorItemTestBuilder().WithColorCodeAndValues("Blue", Color.Blue.ToArgb(), 4.5, 3.5, 1.5).Build(),
                new ColorItemTestBuilder().WithColorCodeAndValues("Yellow", Color.Yellow.ToArgb(), 5.0, 2.5, 1.5)
                    .Build()
            };

            BuildReport(doc, colors, "colorItems", new [] { typeof(ColorItemTestClass) });

            doc.Save(ArtifactsDir + "ReportingEngine.SetPointColorDynamically.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.SetPointColorDynamically.docx", GoldsDir + "ReportingEngine.SetPointColorDynamically Gold.docx"), Is.True);
        }

        [Test]
        public void ConditionalExpressionForLeaveChartSeries()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Chart series.docx");

            int condition = 3;
            BuildReport(doc, new object[] { Common.GetManagers(), condition }, new[] { "managers", "condition" });

            doc.Save(ArtifactsDir + "ReportingEngine.TestLeaveChartSeries.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.TestLeaveChartSeries.docx", GoldsDir + "ReportingEngine.TestLeaveChartSeries Gold.docx"), Is.True);
        }

        [Test, Ignore("WORDSNET-20810")]
        public void ConditionalExpressionForRemoveChartSeries()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Chart series.docx");

            int condition = 2;
            BuildReport(doc, new object[] { Common.GetManagers(), condition }, new[] { "managers", "condition" });

            doc.Save(ArtifactsDir + "ReportingEngine.TestRemoveChartSeries.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.TestRemoveChartSeries.docx", GoldsDir + "ReportingEngine.TestRemoveChartSeries Gold.docx"), Is.True);
        }

        [Test]
        public void IndexOf()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Index of.docx");

            BuildReport(doc, Common.GetManagers(), "Managers");

            doc = DocumentHelper.SaveOpen(doc);

            Assert.That(doc.GetText(), Is.EqualTo("The names are: John Smith, Tony Anderson, July James\f"));
        }

        [Test]
        public void IfElse()
        {
            Document doc = new Document(MyDir + "Reporting engine template - If-else.docx");

            BuildReport(doc, Common.GetManagers(), "m");

            doc = DocumentHelper.SaveOpen(doc);

            Assert.That(doc.GetText(), Is.EqualTo("You have chosen 3 item(s).\f"));
        }

        [Test]
        public void IfElseWithoutData()
        {
            Document doc = new Document(MyDir + "Reporting engine template - If-else.docx");

            BuildReport(doc, Common.GetEmptyManagers(), "m");

            doc = DocumentHelper.SaveOpen(doc);

            Assert.That(doc.GetText(), Is.EqualTo("You have chosen no items.\f"));
        }

        [Test]
        public void ExtensionMethods()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Extension methods.docx");

            BuildReport(doc, Common.GetManagers(), "Managers");

            doc.Save(ArtifactsDir + "ReportingEngine.ExtensionMethods.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.ExtensionMethods.docx", GoldsDir + "ReportingEngine.ExtensionMethods Gold.docx"), Is.True);
        }

        [Test]
        public void Operators()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Operators.docx");

            NumericTestClass testData = new NumericTestBuilder().WithValuesAndLogical(1, 2.0, 3, null, true).Build();

            ReportingEngine report = new ReportingEngine();
            report.KnownTypes.Add(typeof(NumericTestBuilder));
            report.BuildReport(doc, testData, "ds");

            doc.Save(ArtifactsDir + "ReportingEngine.Operators.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.Operators.docx", GoldsDir + "ReportingEngine.Operators Gold.docx"), Is.True);
        }

        [Test]
        public void HeaderVariable()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Header variable.docx");

            BuildReport(doc, new DataSet(), "", ReportBuildOptions.UseLegacyHeaderFooterVisiting);

            doc.Save(ArtifactsDir + "ReportingEngine.HeaderVariable.docx");

            Assert.That(doc.FirstSection.Body.FirstParagraph.GetText().Trim(), Is.EqualTo("Value of myHeaderVariable is: I am header variable"));
        }

        [Test]
        public void ContextualObjectMemberAccess()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Contextual object member access.docx");

            BuildReport(doc, Common.GetManagers(), "Managers");

            doc.Save(ArtifactsDir + "ReportingEngine.ContextualObjectMemberAccess.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.ContextualObjectMemberAccess.docx", GoldsDir + "ReportingEngine.ContextualObjectMemberAccess Gold.docx"), Is.True);
        }

        [Test]
        public void InsertDocumentDynamicallyWithAdditionalTemplateChecking()
        {
            Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.Document] -build>>");

            DocumentTestClass doc = new DocumentTestBuilder()
                .WithDocument(new Document(MyDir + "Reporting engine template - Data table.docx")).Build();

            BuildReport(template, new object[] { doc, Common.GetContracts() }, new[] { "src", "Contracts" }, 
                ReportBuildOptions.None);
            template.Save(
                ArtifactsDir + "ReportingEngine.InsertDocumentDynamicallyWithAdditionalTemplateChecking.docx");

            Assert.That(DocumentHelper.CompareDocs(
                    ArtifactsDir + "ReportingEngine.InsertDocumentDynamicallyWithAdditionalTemplateChecking.docx",
                    GoldsDir + "ReportingEngine.InsertDocumentDynamicallyWithAdditionalTemplateChecking Gold.docx"), Is.True, "Fail inserting document by document");
        }

        [Test]
        public void InsertDocumentDynamicallyWithStyles()
        {
            Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.Document] -sourceStyles>>");

            DocumentTestClass doc = new DocumentTestBuilder()
                .WithDocument(new Document(MyDir + "Reporting engine template - Data table.docx")).Build();

            BuildReport(template, doc, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx", GoldsDir + "ReportingEngine.InsertDocumentDynamically(stream,doc,bytes) Gold.docx"), Is.True, "Fail inserting document by document");
        }

        [Test]
        public void InsertDocumentDynamicallyTrimLastParagraph()
        {
            Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.Document] -inline>>");

            DocumentTestClass doc = new DocumentTestBuilder()
                .WithDocument(new Document(MyDir + "Reporting engine template - Data table.docx")).Build();

            BuildReport(template, doc, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx");

            template = new Document(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx");
            Assert.That(template.FirstSection.Body.Paragraphs.Count, Is.EqualTo(1));
        }

        [Test]
        public void SourseListNumbering()
        {
            //ExStart:SourseListNumbering
            //GistId:6e4482e7434754c31c6f2f6e4bf48bb1
            //ExFor:ReportingEngine.BuildReport(Document, Object[], String[])
            //ExSummary:Shows how to keep inserted numbering as is.
            // By default, numbered lists from a template document are continued when their identifiers match those from a document being inserted.
            // With "-sourceNumbering" numbering should be separated and kept as is.
            Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.Document]>>" + Environment.NewLine + "<<doc [src.Document] -sourceNumbering>>");

            DocumentTestClass doc = new DocumentTestBuilder()
                .WithDocument(new Document(MyDir + "List item.docx")).Build();

            ReportingEngine engine = new ReportingEngine() { Options = ReportBuildOptions.RemoveEmptyParagraphs };
            engine.BuildReport(template, new object[] { doc }, new[] { "src" });

            template.Save(ArtifactsDir + "ReportingEngine.SourseListNumbering.docx");
            //ExEnd:SourseListNumbering

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.SourseListNumbering.docx", GoldsDir + "ReportingEngine.SourseListNumbering Gold.docx"), Is.True);
        }

        [Test]
        public void InsertDocumentDynamicallyByStream()
        {
            Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.DocumentStream]>>");

            DocumentTestClass docStream = new DocumentTestBuilder()
                .WithDocumentStream(new FileStream(mDocument, FileMode.Open, FileAccess.Read)).Build();

            BuildReport(template, docStream, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx", GoldsDir + "ReportingEngine.InsertDocumentDynamically(stream,doc,bytes) Gold.docx"), Is.True, "Fail inserting document by stream");
        }

        [Test]
        public void InsertDocumentDynamicallyByBytes()
        {
            Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.DocumentBytes]>>");

            DocumentTestClass docBytes = new DocumentTestBuilder()
                .WithDocumentBytes(File.ReadAllBytes(MyDir + "Reporting engine template - Data table.docx")).Build();

            BuildReport(template, docBytes, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx", GoldsDir + "ReportingEngine.InsertDocumentDynamically(stream,doc,bytes) Gold.docx"), Is.True, "Fail inserting document by bytes");
        }

        [Test]
        public void InsertDocumentDynamicallyByUri()
        {
            Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.DocumentString]>>");

            DocumentTestClass docUri = new DocumentTestBuilder()
                .WithDocumentString("http://www.snee.com/xml/xslt/sample.doc").Build();

            BuildReport(template, docUri, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx", GoldsDir + "ReportingEngine.InsertDocumentDynamically(uri) Gold.docx"), Is.True, "Fail inserting document by uri");
        }

        [Test]
        public void InsertDocumentDynamicallyByBase64()
        {
            Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.DocumentString]>>");
            string base64Template = File.ReadAllText(MyDir + "Reporting engine template - Data table (base64).txt");

            DocumentTestClass docBase64 = new DocumentTestBuilder().WithDocumentString(base64Template).Build();

            BuildReport(template, docBase64, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx", GoldsDir + "ReportingEngine.InsertDocumentDynamically(stream,doc,bytes) Gold.docx"), Is.True, "Fail inserting document by uri");
        }

        [Test]
        public void InsertImageDynamically()
        {
            Document template =
                DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.Image]>>", ShapeType.TextBox);

            ImageTestClass image = new ImageTestBuilder().WithImage(mImage).Build();

            BuildReport(template, image, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.InsertImageDynamically.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.InsertImageDynamically.docx", GoldsDir + "ReportingEngine.InsertImageDynamically(stream,doc,bytes) Gold.docx"), Is.True, "Fail inserting document by bytes");
        }

        [Test]
        public void InsertImageDynamicallyByStream()
        {
            Document template =
                DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.ImageStream]>>", ShapeType.TextBox);
            ImageTestClass imageStream = new ImageTestBuilder()
                .WithImageStream(new FileStream(mImage, FileMode.Open, FileAccess.Read)).Build();

            BuildReport(template, imageStream, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.InsertImageDynamically.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.InsertImageDynamically.docx", GoldsDir + "ReportingEngine.InsertImageDynamically(stream,doc,bytes) Gold.docx"), Is.True, "Fail inserting document by bytes");
        }

        [Test]
        public void InsertImageDynamicallyByBytes()
        {
            Document template =
                DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.ImageBytes]>>", ShapeType.TextBox);
            ImageTestClass imageBytes = new ImageTestBuilder().WithImageBytes(File.ReadAllBytes(mImage)).Build();

            BuildReport(template, imageBytes, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.InsertImageDynamically.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.InsertImageDynamically.docx", GoldsDir + "ReportingEngine.InsertImageDynamically(stream,doc,bytes) Gold.docx"), Is.True, "Fail inserting document by bytes");
        }

        [Test]
        public void InsertImageDynamicallyByUri()
        {
            Document template =
                DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.ImageString]>>", ShapeType.TextBox);
            ImageTestClass imageUri = new ImageTestBuilder()
                .WithImageString("https://metrics.aspose.com/img/headergraphics.svg")
                .Build();

            BuildReport(template, imageUri, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.InsertImageDynamically.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.InsertImageDynamically.docx",
                    GoldsDir + "ReportingEngine.InsertImageDynamically(uri) Gold.docx"), Is.True, "Fail inserting document by bytes");
        }

        [Test]
        public void InsertImageDynamicallyByBase64()
        {
            Document template =
                DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.ImageString]>>", ShapeType.TextBox);
            string base64Template = File.ReadAllText(MyDir + "Reporting engine template - base64 image.txt");

            ImageTestClass imageBase64 = new ImageTestBuilder().WithImageString(base64Template).Build();

            BuildReport(template, imageBase64, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.InsertImageDynamically.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.InsertImageDynamically.docx",
                    GoldsDir + "ReportingEngine.InsertImageDynamically(stream,doc,bytes) Gold.docx"), Is.True, "Fail inserting document by bytes");

        }

        [TestCase("<<[html_text] -html>>")]
        [TestCase("<<html [html_text]>>")]
        [TestCase("<<html [html_text] -sourceStyles>>")]
        public void InsertHtmlDinamically(string templateText)
        {
            string html = File.ReadAllText(MyDir + "Reporting engine template - Html.html");

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln(templateText);

            BuildReport(doc, html, "html_text");
            doc.Save(ArtifactsDir + "ReportingEngine.InsertHtmlDinamically.docx");
        }

        [Test]
        public void ImageExifOrientation()
        {
            Document template = new Document(MyDir + "Reporting engine template - Image exif orientation.docx");

            byte[] image1Bytes = File.ReadAllBytes(ImageDir + "RightF.jpg");
            byte[] image2Bytes = File.ReadAllBytes(ImageDir + "WrongF.jpg");

            BuildReport(template, new object[] { image1Bytes, image2Bytes }, new string[] { "image1", "image2" }, 
                ReportBuildOptions.RespectJpegExifOrientation);
            template.Save(ArtifactsDir + "ReportingEngine.ImageExifOrientation.docx");
        }

        [Test]
        public void DynamicStretchingImageWithinTextBox()
        {
            Document template = new Document(MyDir + "Reporting engine template - Dynamic stretching.docx");
            
            ImageTestClass image = new ImageTestBuilder().WithImage(mImage).Build();

            BuildReport(template, image, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.DynamicStretchingImageWithinTextBox.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.DynamicStretchingImageWithinTextBox.docx",
                    GoldsDir + "ReportingEngine.DynamicStretchingImageWithinTextBox Gold.docx"), Is.True);
        }

        [TestCase("https://auckland.dynabic.com/wiki/display/org/Supported+dynamic+insertion+of+hyperlinks+for+LINQ+Reporting+Engine")]
        [TestCase("Bookmark")]
        public void InsertHyperlinksDynamically(string link)
        {
            Document template = new Document(MyDir + "Reporting engine template - Inserting hyperlinks.docx");
            BuildReport(template, 
                new object[]
                {
                    link, // Use URI or the name of a bookmark within the same document for a hyperlink
                    "Aspose"
                },
                new[]
                {
                    "uri_or_bookmark_expression", 
                    "display_text_expression"
                });

            template.Save(ArtifactsDir + "ReportingEngine.InsertHyperlinksDynamically.docx");
        }

        [Test]
        public void InsertBookmarksDynamically()
        {
            Document doc =
                DocumentHelper.CreateSimpleDocument(
                    "<<bookmark [bookmark_expression]>><<foreach [m in Contracts]>><<[m.Client.Name]>><</foreach>><</bookmark>>");

            BuildReport(doc, new object[] { "BookmarkOne", Common.GetContracts() },
                new[] { "bookmark_expression", "Contracts" });

            doc.Save(ArtifactsDir + "ReportingEngine.InsertBookmarksDynamically.docx");
        }

        [Test]
        public void WithoutKnownType()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("<<[new DateTime()]:”dd.MM.yyyy”>>");

            ReportingEngine engine = new ReportingEngine();
            Assert.Throws<InvalidOperationException>(() => engine.BuildReport(doc, ""));
        }

        [Test]
        public void WorkWithKnownTypes()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("<<[new DateTime(2016, 1, 20)]:”dd.MM.yyyy”>>");
            builder.Writeln("<<[new DateTime(2016, 1, 20)]:”dd”>>");
            builder.Writeln("<<[new DateTime(2016, 1, 20)]:”MM”>>");
            builder.Writeln("<<[new DateTime(2016, 1, 20)]:”yyyy”>>");
            builder.Writeln("<<[new DateTime(2016, 1, 20).Month]>>");

            BuildReport(doc, "", new []{ typeof(DateTime) });

            doc.Save(ArtifactsDir + "ReportingEngine.KnownTypes.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.KnownTypes.docx", GoldsDir + "ReportingEngine.KnownTypes Gold.docx"), Is.True);
        }

        [Test]
        public void WorkWithContentControls()
        {
            Document doc = new Document(MyDir + "Reporting engine template - CheckBox Content Control.docx");
            BuildReport(doc, Common.GetManagers(), "Managers");

            doc.Save(ArtifactsDir + "ReportingEngine.WorkWithContentControls.docx");
        }

        [Test]
        public void WorkWithSingleColumnTableRow()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Table row.docx");
            BuildReport(doc, Common.GetManagers(), "Managers");

            doc.Save(ArtifactsDir + "ReportingEngine.SingleColumnTableRow.docx");
        }

        [Test]
        public void WorkWithSingleColumnTableRowGreedy()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Table row greedy.docx");
            BuildReport(doc, Common.GetManagers(), "Managers");

            doc.Save(ArtifactsDir + "ReportingEngine.SingleColumnTableRowGreedy.docx");
        }

        [Test]
        public void TableRowConditionalBlocks()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Table row conditional blocks.docx");

            List<ClientTestClass> clients = new List<ClientTestClass>
            {
                new ClientTestClass
                {
                    Name = "John Monrou",
                    Country = "France",
                    LocalAddress = "27 RUE PASTEUR"
                },
                new ClientTestClass
                {
                    Name = "James White",
                    Country = "England",
                    LocalAddress = "14 Tottenham Court Road"
                },
                new ClientTestClass
                {
                    Name = "Kate Otts",
                    Country = "New Zealand",
                    LocalAddress = "Wellington 6004"
                }
            };

            BuildReport(doc, clients, "clients");

            doc.Save(ArtifactsDir + "ReportingEngine.TableRowConditionalBlocks.docx");
        }

        [Test]
        public void IfGreedy()
        {
            Document doc = new Document(MyDir + "Reporting engine template - If greedy.docx");

            AsposeData obj = new AsposeData
            {
                List = new List<string>
                {
                    "abc"
                }
            };

            BuildReport(doc, obj);

            doc.Save(ArtifactsDir + "ReportingEngine.IfGreedy.docx");
        }

        public class AsposeData
        {
            public List<string> List { get; set; }
        }

        [Test]
        public void StretchImagefitHeight()
        {
            Document doc =
                DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.ImageStream] -fitHeight>>",
                    ShapeType.TextBox);

            ImageTestClass imageStream = new ImageTestBuilder()
                .WithImageStream(new FileStream(mImage, FileMode.Open, FileAccess.Read)).Build();
            BuildReport(doc, imageStream, "src", ReportBuildOptions.None);

            doc = DocumentHelper.SaveOpen(doc);

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes.OfType<Shape>())
            {
                // Assert that the image is really insert in textbox.
                Assert.That(shape.Fill.ImageBytes, Is.Not.Null);

                // Assert that the width is preserved, and the height is changed.
                Assert.That(shape.Height, Is.Not.EqualTo(346.35));
                Assert.That(shape.Width, Is.EqualTo(431.5));
            }
        }

        [Test]
        public void StretchImagefitWidth()
        {
            Document doc =
                DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.ImageStream] -fitWidth>>",
                    ShapeType.TextBox);

            ImageTestClass imageStream = new ImageTestBuilder()
                .WithImageStream(new FileStream(mImage, FileMode.Open, FileAccess.Read)).Build();
            BuildReport(doc, imageStream, "src", ReportBuildOptions.None);

            doc = DocumentHelper.SaveOpen(doc);

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes.OfType<Shape>())
            {
                Assert.That(shape.Fill.ImageBytes, Is.Not.Null);

                // Assert that the height is preserved, and the width is changed.
                Assert.That(shape.Width, Is.Not.EqualTo(431.5));
                Assert.That(shape.Height, Is.EqualTo(346.35));
            }
        }

        [Test]
        public void StretchImagefitSize()
        {
            Document doc =
                DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.ImageStream] -fitSize>>",
                    ShapeType.TextBox);

            ImageTestClass imageStream = new ImageTestBuilder()
                .WithImageStream(new FileStream(mImage, FileMode.Open, FileAccess.Read)).Build();
            BuildReport(doc, imageStream, "src", ReportBuildOptions.None);

            doc = DocumentHelper.SaveOpen(doc);

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes.OfType<Shape>())
            {
                Assert.That(shape.Fill.ImageBytes, Is.Not.Null);
                
                // Assert that the height and the width are changed.
                Assert.That(shape.Height, Is.Not.EqualTo(346.35));
                Assert.That(shape.Width, Is.Not.EqualTo(431.5));
            }
        }

        [Test]
        public void StretchImagefitSizeLim()
        {
            Document doc =
                DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.ImageStream] -fitSizeLim>>",
                    ShapeType.TextBox);

            ImageTestClass imageStream = new ImageTestBuilder()
                .WithImageStream(new FileStream(mImage, FileMode.Open, FileAccess.Read)).Build();
            BuildReport(doc, imageStream, "src", ReportBuildOptions.None);

            doc = DocumentHelper.SaveOpen(doc);

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes.OfType<Shape>())
            {
                Assert.That(shape.Fill.ImageBytes, Is.Not.Null);

                // Assert that textbox size are equal image size.
                Assert.That(shape.Height, Is.EqualTo(300.0d));
                Assert.That(shape.Width, Is.EqualTo(300.0d));
            }
        }

        [Test]
        public void WithoutMissingMembers()
        {
            DocumentBuilder builder = new DocumentBuilder();

            // Add templete to the document for reporting engine.
            DocumentHelper.InsertBuilderText(builder,
                new[] { "<<[missingObject.First().id]>>", "<<foreach [in missingObject]>><<[id]>><</foreach>>" });

            // Assert that build report failed without "ReportBuildOptions.AllowMissingMembers".
            Assert.Throws<InvalidOperationException>(
                () => BuildReport(builder.Document, new DataSet(), "", ReportBuildOptions.None));
        }

        [Test]
        public void MissingMembers()
        {
            //ExStart:MissingMembers
            //GistId:65919861586e42e24f61a3ccb65f8f4e
            //ExFor:ReportingEngine.BuildReport(Document, Object, String)
            //ExFor:ReportingEngine.MissingMemberMessage
            //ExFor:ReportingEngine.Options
            //ExSummary:Shows how to allow missinng members.
            DocumentBuilder builder = new DocumentBuilder();
            builder.Writeln("<<[missingObject.First().id]>>");
            builder.Writeln("<<foreach [in missingObject]>><<[id]>><</foreach>>");

            ReportingEngine engine = new ReportingEngine { Options = ReportBuildOptions.AllowMissingMembers };
            engine.MissingMemberMessage = "Missed";
            engine.BuildReport(builder.Document, new DataSet(), "");
            //ExEnd:MissingMembers

            // Assert that build report success with "ReportBuildOptions.AllowMissingMembers".
            Assert.That(builder.Document.GetText().Trim(), Is.EqualTo("Missed"));
        }

        [TestCase("<<[missingObject.First().id]>>", "<<[missingObject.First( Error! Can not get the value of member 'missingObject' on type 'System.Data.DataSet'. ).id]>>", TestName = "Can not get the value of member")]
        [TestCase("<<[new DateTime()]:\"dd.MM.yyyy\">>", "<<[new DateTime( Error! A type identifier is expected. )]:\"dd.MM.yyyy\">>", TestName = "A type identifier is expected")]
        [TestCase("<<]>>", "<<] Error! Character ']' is unexpected. >>", TestName = "Character is unexpected")]
        [TestCase("<<[>>", "<<[>> Error! An expression is expected.", TestName = "An expression is expected")]
        [TestCase("<<>>", "<<>> Error! Tag end is unexpected.", TestName = "Tag end is unexpected")]
        public void InlineErrorMessages(string templateText, string result)
        {
            DocumentBuilder builder = new DocumentBuilder();
            DocumentHelper.InsertBuilderText(builder, new[] { templateText });

            BuildReport(builder.Document, new DataSet(), "", ReportBuildOptions.InlineErrorMessages);

            Assert.That(builder.Document.FirstSection.Body.Paragraphs[0].GetText().TrimEnd(), Is.EqualTo(result));
        }

        [Test]
        public void SetBackgroundColorDynamically()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Background color.docx");

            List<ColorItemTestClass> colors = new List<ColorItemTestClass>
            {
                new ColorItemTestBuilder().WithColor("Black", Color.Black).Build(),
                new ColorItemTestBuilder().WithColor("Red", Color.FromArgb(255, 0, 0)).Build(),
                new ColorItemTestBuilder().WithColor("Empty", Color.Empty).Build()
            };

            BuildReport(doc, colors, "Colors");

            doc.Save(ArtifactsDir + "ReportingEngine.SetBackgroundColorDynamically.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.SetBackgroundColorDynamically.docx",
                GoldsDir + "ReportingEngine.SetBackgroundColorDynamically Gold.docx"), Is.True);
        }

        [Test]
        public void SetTextColorDynamically()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Text color.docx");

            List<ColorItemTestClass> colors = new List<ColorItemTestClass>
            {
                new ColorItemTestBuilder().WithColor("Black", Color.Blue).Build(),
                new ColorItemTestBuilder().WithColor("Red", Color.FromArgb(255, 0, 0)).Build(),
                new ColorItemTestBuilder().WithColor("Empty", Color.Empty).Build()
            };

            BuildReport(doc, colors, "Colors");

            doc.Save(ArtifactsDir + "ReportingEngine.SetTextColorDynamically.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.SetTextColorDynamically.docx",
                GoldsDir + "ReportingEngine.SetTextColorDynamically Gold.docx"), Is.True);
        }

        [Test]
        public void DoNotRemoveEmptyParagraphs()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Remove empty paragraphs.docx");

            BuildReport(doc, Common.GetManagers(), "Managers");

            doc.Save(ArtifactsDir + "ReportingEngine.DoNotRemoveEmptyParagraphs.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.DoNotRemoveEmptyParagraphs.docx",
                GoldsDir + "ReportingEngine.DoNotRemoveEmptyParagraphs Gold.docx"), Is.True);
        }

        [Test]
        public void RemoveEmptyParagraphs()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Remove empty paragraphs.docx");

            BuildReport(doc, Common.GetManagers(), "Managers", ReportBuildOptions.RemoveEmptyParagraphs);

            doc.Save(ArtifactsDir + "ReportingEngine.RemoveEmptyParagraphs.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.RemoveEmptyParagraphs.docx",
                GoldsDir + "ReportingEngine.RemoveEmptyParagraphs Gold.docx"), Is.True);
        }

        [TestCase("Hello", "Hello", "ReportingEngine.MergingTableCellsDynamically.Merged", TestName = "Cells in the first two tables must be merged")]
        [TestCase("Hello", "Name", "ReportingEngine.MergingTableCellsDynamically.NotMerged", TestName = "Only last table cells must be merge")]
        public void MergingTableCellsDynamically(string value1, string value2, string resultDocumentName)
        {
            string artifactPath = ArtifactsDir + resultDocumentName +
                                   FileFormatUtil.SaveFormatToExtension(SaveFormat.Docx);
            string goldPath = GoldsDir + resultDocumentName + " Gold" +
                              FileFormatUtil.SaveFormatToExtension(SaveFormat.Docx);

            Document doc = new Document(MyDir + "Reporting engine template - Merging table cells dynamically.docx");

            List<ClientTestClass> clients = new List<ClientTestClass>
            {
                new ClientTestClass
                {
                    Name = "John Monrou",
                    Country = "France",
                    LocalAddress = "27 RUE PASTEUR"
                },
                new ClientTestClass
                {
                    Name = "James White",
                    Country = "New Zealand",
                    LocalAddress = "14 Tottenham Court Road"
                },
                new ClientTestClass
                {
                    Name = "Kate Otts",
                    Country = "New Zealand",
                    LocalAddress = "Wellington 6004"
                }
            };

            BuildReport(doc, new object[] { value1, value2, clients }, new [] { "value1", "value2", "clients" });
            doc.Save(artifactPath);

            Assert.That(DocumentHelper.CompareDocs(artifactPath, goldPath), Is.True);
        }

        [Test]
        public void XmlDataStringWithoutSchema()
        {
            //ExStart
            //ExFor:XmlDataSource
            //ExFor:XmlDataSource.#ctor(String)
            //ExSummary:Show how to use XML as a data source (string).
            Document doc = new Document(MyDir + "Reporting engine template - XML data destination.docx");

            XmlDataSource dataSource = new XmlDataSource(MyDir + "List of people.xml");
            BuildReport(doc, dataSource, "persons");

            doc.Save(ArtifactsDir + "ReportingEngine.XmlDataString.docx");
            //ExEnd

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.XmlDataString.docx",
                GoldsDir + "ReportingEngine.DataSource Gold.docx"), Is.True);
        }

        [Test]
        public void XmlDataStreamWithoutSchema()
        {
            //ExStart
            //ExFor:XmlDataSource
            //ExFor:XmlDataSource.#ctor(Stream)
            //ExSummary:Show how to use XML as a data source (stream).
            Document doc = new Document(MyDir + "Reporting engine template - XML data destination.docx");

            using (FileStream stream = File.OpenRead(MyDir + "List of people.xml"))
            {
                XmlDataSource dataSource = new XmlDataSource(stream);
                BuildReport(doc, dataSource, "persons");
            }

            doc.Save(ArtifactsDir + "ReportingEngine.XmlDataStream.docx");
            //ExEnd

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.XmlDataStream.docx",
                GoldsDir + "ReportingEngine.DataSource Gold.docx"), Is.True);
        }

        [Test]
        public void XmlDataWithNestedElements()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Data destination with nested elements.docx");

            XmlDataSource dataSource = new XmlDataSource(MyDir + "Nested elements.xml");
            BuildReport(doc, dataSource, "managers");

            doc.Save(ArtifactsDir + "ReportingEngine.XmlDataWithNestedElements.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.XmlDataWithNestedElements.docx",
                GoldsDir + "ReportingEngine.DataSourceWithNestedElements Gold.docx"), Is.True);
        }

        [Test]
        public void JsonDataString()
        {
            //ExStart
            //ExFor:JsonDataLoadOptions
            //ExFor:JsonDataLoadOptions.#ctor
            //ExFor:JsonDataLoadOptions.ExactDateTimeParseFormats
            //ExFor:JsonDataLoadOptions.AlwaysGenerateRootObject
            //ExFor:JsonDataLoadOptions.PreserveSpaces
            //ExFor:JsonDataLoadOptions.SimpleValueParseMode
            //ExFor:JsonDataSource
            //ExFor:JsonDataSource.#ctor(String,JsonDataLoadOptions)
            //ExFor:JsonSimpleValueParseMode
            //ExSummary:Shows how to use JSON as a data source (string).
            Document doc = new Document(MyDir + "Reporting engine template - JSON data destination.docx");

            JsonDataLoadOptions options = new JsonDataLoadOptions
            {
                ExactDateTimeParseFormats = new List<string> {"MM/dd/yyyy", "MM.d.yy", "MM d yy"},
                AlwaysGenerateRootObject = true,
                PreserveSpaces = true,
                SimpleValueParseMode = JsonSimpleValueParseMode.Loose
            };

            JsonDataSource dataSource = new JsonDataSource(MyDir + "List of people.json", options);
            BuildReport(doc, dataSource, "persons");

            doc.Save(ArtifactsDir + "ReportingEngine.JsonDataString.docx");
            //ExEnd

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.JsonDataString.docx",
                GoldsDir + "ReportingEngine.JsonDataString Gold.docx"), Is.True);
        }

        [Test]
        public void JsonDataStringException()
        {
            Document doc = new Document(MyDir + "Reporting engine template - JSON data destination.docx");

            JsonDataLoadOptions options = new JsonDataLoadOptions();
            options.SimpleValueParseMode = JsonSimpleValueParseMode.Strict;

            JsonDataSource dataSource = new JsonDataSource(MyDir + "List of people.json", options);
            Assert.Throws<InvalidOperationException>(() => BuildReport(doc, dataSource, "persons"));
        }

        [Test]
        public void JsonDataStream()
        {
            //ExStart
            //ExFor:JsonDataSource.#ctor(Stream,JsonDataLoadOptions)
            //ExSummary:Shows how to use JSON as a data source (stream).
            Document doc = new Document(MyDir + "Reporting engine template - JSON data destination.docx");

            JsonDataLoadOptions options = new JsonDataLoadOptions
            {
                ExactDateTimeParseFormats = new List<string> {"MM/dd/yyyy", "MM.d.yy", "MM d yy"}
            };

            using (FileStream stream = File.OpenRead(MyDir + "List of people.json"))
            {
                JsonDataSource dataSource = new JsonDataSource(stream, options);
                BuildReport(doc, dataSource, "persons");
            }

            doc.Save(ArtifactsDir + "ReportingEngine.JsonDataStream.docx");
            //ExEnd

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.JsonDataStream.docx",
                GoldsDir + "ReportingEngine.JsonDataString Gold.docx"), Is.True);
        }

        [Test]
        public void JsonDataWithNestedElements()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Data destination with nested elements.docx");

            JsonDataSource dataSource = new JsonDataSource(MyDir + "Nested elements.json");
            BuildReport(doc, dataSource, "managers");

            doc.Save(ArtifactsDir + "ReportingEngine.JsonDataWithNestedElements.docx");

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.JsonDataWithNestedElements.docx",
                GoldsDir + "ReportingEngine.DataSourceWithNestedElements Gold.docx"), Is.True);
        }

        [Test]
        public void JsonDataPreserveSpaces()
        {
            const string template = "LINE BEFORE\r<<[LineWhitespace]>>\r<<[BlockWhitespace]>>LINE AFTER";
            const string expectedResult = "LINE BEFORE\r    \r\r\r\r\rLINE AFTER";
            const string json =
                "{" +
                "    \"LineWhitespace\" : \"    \"," +
                "    \"BlockWhitespace\" : \"\r\n\r\n\r\n\r\n\"" +
                "}";

            using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                JsonDataLoadOptions options = new JsonDataLoadOptions();
                options.PreserveSpaces = true;
                options.SimpleValueParseMode = JsonSimpleValueParseMode.Strict;

                JsonDataSource dataSource = new JsonDataSource(stream, options);

                DocumentBuilder builder = new DocumentBuilder();
                builder.Write(template);

                BuildReport(builder.Document, dataSource, "ds");

                Assert.That(builder.Document.GetText(), Is.EqualTo(expectedResult + ControlChar.SectionBreak));
            }
        }

        [Test]
        public void CsvDataString()
        {
            //ExStart
            //ExFor:CsvDataLoadOptions
            //ExFor:CsvDataLoadOptions.#ctor
            //ExFor:CsvDataLoadOptions.#ctor(Boolean)
            //ExFor:CsvDataLoadOptions.Delimiter
            //ExFor:CsvDataLoadOptions.CommentChar
            //ExFor:CsvDataLoadOptions.HasHeaders
            //ExFor:CsvDataLoadOptions.QuoteChar
            //ExFor:CsvDataSource
            //ExFor:CsvDataSource.#ctor(String,CsvDataLoadOptions)
            //ExSummary:Shows how to use CSV as a data source (string).
            Document doc = new Document(MyDir + "Reporting engine template - CSV data destination.docx");

            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
            loadOptions.Delimiter = ';';
            loadOptions.CommentChar = '$';
            loadOptions.HasHeaders = true;
            loadOptions.QuoteChar = '"';

            CsvDataSource dataSource = new CsvDataSource(MyDir + "List of people.csv", loadOptions);
            BuildReport(doc, dataSource, "persons");

            doc.Save(ArtifactsDir + "ReportingEngine.CsvDataString.docx");
            //ExEnd

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.CsvDataString.docx",
                GoldsDir + "ReportingEngine.CsvData Gold.docx"), Is.True);
        }

        [Test]
        public void CsvDataStream()
        {
            //ExStart
            //ExFor:CsvDataSource.#ctor(Stream,CsvDataLoadOptions)
            //ExSummary:Shows how to use CSV as a data source (stream).
            Document doc = new Document(MyDir + "Reporting engine template - CSV data destination.docx");

            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
            loadOptions.Delimiter = ';';
            loadOptions.CommentChar = '$';

            using (FileStream stream = File.OpenRead(MyDir + "List of people.csv"))
            {
                CsvDataSource dataSource = new CsvDataSource(stream, loadOptions);
                BuildReport(doc, dataSource, "persons");
            }

            doc.Save(ArtifactsDir + "ReportingEngine.CsvDataStream.docx");
            //ExEnd

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.CsvDataStream.docx",
                GoldsDir + "ReportingEngine.CsvData Gold.docx"), Is.True);
        }

        [TestCase(SdtType.ComboBox)]
        [TestCase(SdtType.DropDownList)]
        public void InsertComboboxDropdownListItemsDynamically(SdtType sdtType)
        {
            const string template =
                "<<item[\"three\"] [\"3\"]>><<if [false]>><<item [\"four\"] [null]>><</if>><<item[\"five\"] [\"5\"]>>";

            SdtListItem[] staticItems =
            {
                new SdtListItem("1", "one"),
                new SdtListItem("2", "two")
            };

            Document doc = new Document();

            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, sdtType, MarkupLevel.Block) { Title = template };

            foreach (SdtListItem item in staticItems)
            {
                sdt.ListItems.Add(item);
            }

            doc.FirstSection.Body.AppendChild(sdt);

            BuildReport(doc, new object(), "");

            doc.Save(ArtifactsDir + $"ReportingEngine.InsertComboboxDropdownListItemsDynamically_{sdtType}.docx");

            doc = new Document(ArtifactsDir +
                               $"ReportingEngine.InsertComboboxDropdownListItemsDynamically_{sdtType}.docx");

            sdt = (StructuredDocumentTag) doc.GetChild(NodeType.StructuredDocumentTag, 0, true);

            SdtListItem[] expectedItems =
            {
                new SdtListItem("1", "one"),
                new SdtListItem("2", "two"),
                new SdtListItem("3", "three"),
                new SdtListItem("5", "five")
            };

            Assert.That(sdt.ListItems.Count, Is.EqualTo(expectedItems.Length));

            for (int i = 0; i < expectedItems.Length; i++)
            {
                Assert.That(sdt.ListItems[i].Value, Is.EqualTo(expectedItems[i].Value));
                Assert.That(sdt.ListItems[i].DisplayText, Is.EqualTo(expectedItems[i].DisplayText));
            }
        }

        [Test]
        public void UpdateFieldsSyntaxAware()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Fields.docx");

            // Note that enabling of the option makes the engine to update fields while building a report,
            // so there is no need to update fields separately after that.            
            BuildReport(doc, new string[] { "First topic", "Second topic", "Third topic" }, "topics",
                ReportBuildOptions.UpdateFieldsSyntaxAware);

            doc.Save(ArtifactsDir + "ReportingEngine.UpdateFieldsSyntaxAware.docx");
        }

        [Test]
        public void DollarTextFormat()
        {
            //ExStart:DollarTextFormat
            //GistId:e386727403c2341ce4018bca370a5b41
            //ExFor:ReportingEngine.BuildReport(Document, Object, String)
            //ExSummary:Shows how to display values as dollar text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("<<[ds.Value1]:dollarText>>\r<<[ds.Value2]:dollarText>>");

            NumericTestClass testData = new NumericTestBuilder().WithValues(1234, 5621718.589).Build();

            ReportingEngine report = new ReportingEngine();
            report.KnownTypes.Add(typeof(NumericTestClass));
            report.BuildReport(doc, testData, "ds");

            doc.Save(ArtifactsDir + "ReportingEngine.DollarTextFormat.docx");
            //ExEnd:DollarTextFormat

            Assert.That(doc.GetText(), Is.EqualTo("one thousand two hundred thirty-four and 00/100\rfive million six hundred twenty-one thousand seven hundred eighteen and 59/100\r\f"));
        }

        [Test,Ignore("To avoid exception with 'SetRestrictedTypes' after execution other tests.")]
        public void RestrictedTypes()
        {
            //ExStart:RestrictedTypes
            //GistId:eeeec1fbf118e95e7df3f346c91ed726
            //ExFor:ReportingEngine.SetRestrictedTypes(Type[])
            //ExSummary:Shows how to deny access to members of types considered insecure.
            Document doc =
                DocumentHelper.CreateSimpleDocument(
                    "<<var [typeVar = \"\".GetType().BaseType]>><<[typeVar]>>");

            // Note, that you can't set restricted types during or after building a report.
            ReportingEngine.SetRestrictedTypes(typeof(System.Type));
            // We set "AllowMissingMembers" option to avoid exceptions during building a report.
            ReportingEngine engine = new ReportingEngine() { Options = ReportBuildOptions.AllowMissingMembers };
            engine.BuildReport(doc, new object());

            // We get an empty string because we can't access the GetType() method.
            Assert.That(doc.GetText().Trim(), Is.EqualTo(string.Empty));
            //ExEnd:RestrictedTypes
        }

        [Test]
        public void Word2016Charts()
        {
            //ExStart:Word2016Charts
            //GistId:a775441ecb396eea917a2717cb9e8f8f
            //ExFor:ReportingEngine.BuildReport(Document, Object[], String[])
            //ExSummary:Shows how to work with charts from word 2016.
            Document doc = new Document(MyDir + "Reporting engine template - Word 2016 Charts.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, new object[] { Common.GetShares(), Common.GetShareQuotes() },
                new string[] { "shares", "quotes" });

            doc.Save(ArtifactsDir + "ReportingEngine.Word2016Charts.docx");
            //ExEnd:Word2016Charts
        }

        [Test]
        public void RemoveParagraphsSelectively()
        {
            //ExStart:RemoveParagraphsSelectively
            //GistId:65919861586e42e24f61a3ccb65f8f4e
            //ExFor:ReportingEngine.BuildReport(Document, Object, String)
            //ExSummary:Shows how to remove paragraphs selectively.
            // Template contains tags with an exclamation mark. For such tags, empty paragraphs will be removed.
            Document doc = new Document(MyDir + "Reporting engine template - Selective remove paragraphs.docx");

            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, false, "value");

            doc.Save(ArtifactsDir + "ReportingEngine.SelectiveDeletionOfParagraphs.docx");
            //ExEnd:RemoveParagraphsSelectively

            Assert.That(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.SelectiveDeletionOfParagraphs.docx", GoldsDir + "ReportingEngine.SelectiveDeletionOfParagraphs Gold.docx"), Is.True);
        }

        private static void BuildReport(Document document, object dataSource, ReportBuildOptions reportBuildOptions)
        {
            ReportingEngine engine = new ReportingEngine { Options = reportBuildOptions };
            engine.BuildReport(document, dataSource);
        }

        private static void BuildReport(Document document, object dataSource, string dataSourceName,
            ReportBuildOptions reportBuildOptions)
        {
            ReportingEngine engine = new ReportingEngine { Options = reportBuildOptions };
            engine.BuildReport(document, dataSource, dataSourceName);
        }

        private static void BuildReport(Document document, object[] dataSource, string[] dataSourceName)
        {
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(document, dataSource, dataSourceName);
        }

        private static void BuildReport(Document document, object[] dataSource, string[] dataSourceName,
            ReportBuildOptions reportBuildOptions)
        {
            ReportingEngine engine = new ReportingEngine { Options = reportBuildOptions };
            engine.BuildReport(document, dataSource, dataSourceName);
        }

        private static void BuildReport(Document document, object dataSource, string dataSourceName, Type[] knownTypes)
        {
            ReportingEngine engine = new ReportingEngine();

            foreach (Type knownType in knownTypes)
            {
                engine.KnownTypes.Add(knownType);
            }

            engine.BuildReport(document, dataSource, dataSourceName);
        }

        private static void BuildReport(Document document, object dataSource)
        {
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(document, dataSource);
        }

        private static void BuildReport(Document document, object dataSource, string dataSourceName)
        {
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(document, dataSource, dataSourceName);
        }

        private static void BuildReport(Document document, object dataSource, Type[] knownTypes)
        {
            ReportingEngine engine = new ReportingEngine();

            foreach (Type knownType in knownTypes)
            {
                engine.KnownTypes.Add(knownType);
            }

            engine.BuildReport(document, dataSource);
        }
    }
}

