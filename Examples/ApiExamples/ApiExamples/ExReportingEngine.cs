// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
using ApiExamples.TestData;
using ApiExamples.TestData.TestBuilders;
using ApiExamples.TestData.TestClasses;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Markup;
using Aspose.Words.Reporting;
using NUnit.Framework;
#if NETCOREAPP2_1 || __MOBILE__
using SkiaSharp;
#endif

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

            Assert.AreEqual("LINQ Reporting Engine says: Hello World\f", doc.GetText());
        }

        [Test]
        public void StringFormat()
        {
            Document doc = DocumentHelper.CreateSimpleDocument(
                "<<[s.Name]:lower>> says: <<[s.Message]:upper>>, <<[s.Message]:caps>>, <<[s.Message]:firstCap>>");

            MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "hello world");
            BuildReport(doc, sender, "s");

            doc = DocumentHelper.SaveOpen(doc);

            Assert.AreEqual("linq reporting engine says: HELLO WORLD, Hello World, Hello world\f", doc.GetText());
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

            Assert.AreEqual("A : ii, 200th, FIRST, Two, C8, - 200 -\f", doc.GetText());
        }

        [Test]
        public void TestDataTable()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Data table.docx");

            BuildReport(doc, Common.GetContracts(), "Contracts");

            doc.Save(ArtifactsDir + "ReportingEngine.TestDataTable.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.TestDataTable.docx", GoldsDir + "ReportingEngine.TestDataTable Gold.docx"));
        }

        [Test]
        public void Total()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Total.docx");

            BuildReport(doc, Common.GetContracts(), "Contracts");

            doc.Save(ArtifactsDir + "ReportingEngine.Total.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.Total.docx", GoldsDir + "ReportingEngine.Total Gold.docx"));
        }

        [Test]
        public void TestNestedDataTable()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Nested data table.docx");

            BuildReport(doc, Common.GetManagers(), "Managers");

            doc.Save(ArtifactsDir + "ReportingEngine.TestNestedDataTable.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.TestNestedDataTable.docx", GoldsDir + "ReportingEngine.TestNestedDataTable Gold.docx"));
        }

        [Test]
        public void RestartingListNumberingDynamically()
        {
            Document template = new Document(MyDir + "Reporting engine template - List numbering.docx");

            BuildReport(template, Common.GetManagers(), "Managers", ReportBuildOptions.RemoveEmptyParagraphs);

            template.Save(ArtifactsDir + "ReportingEngine.RestartingListNumberingDynamically.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.RestartingListNumberingDynamically.docx", GoldsDir + "ReportingEngine.RestartingListNumberingDynamically Gold.docx"));
        }

        [Test]
        public void RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically()
        {
            Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.Document] -build>>");
            
            DocumentTestClass doc = new DocumentTestBuilder()
                .WithDocument(new Document(MyDir + "Reporting engine template - List numbering.docx")).Build();

            BuildReport(template, new object[] {doc, Common.GetManagers()} , new[] {"src", "Managers"}, ReportBuildOptions.RemoveEmptyParagraphs);

            template.Save(ArtifactsDir + "ReportingEngine.RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically.docx", GoldsDir + "ReportingEngine.RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically Gold.docx"));
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

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.RestartingListNumberingDynamicallyWhileMultipleInsertionsDocumentDynamically.docx", GoldsDir + "ReportingEngine.RestartingListNumberingDynamicallyWhileInsertingDocumentDynamically Gold.docx"));
         }

        [Test]
        public void ChartTest()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Chart.docx");

            BuildReport(doc, Common.GetManagers(), "managers");

            doc.Save(ArtifactsDir + "ReportingEngine.TestChart.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.TestChart.docx", GoldsDir + "ReportingEngine.TestChart Gold.docx"));
        }

        [Test]
        public void BubbleChartTest()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Bubble chart.docx");

            BuildReport(doc, Common.GetManagers(), "managers");

            doc.Save(ArtifactsDir + "ReportingEngine.TestBubbleChart.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.TestBubbleChart.docx", GoldsDir + "ReportingEngine.TestBubbleChart Gold.docx"));
        }

        [Test]
        public void SetChartSeriesColorsDynamically()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Chart series color.docx");

            BuildReport(doc, Common.GetManagers(), "managers");

            doc.Save(ArtifactsDir + "ReportingEngine.SetChartSeriesColorDynamically.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.SetChartSeriesColorDynamically.docx", GoldsDir + "ReportingEngine.SetChartSeriesColorDynamically Gold.docx"));
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

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.SetPointColorDynamically.docx", GoldsDir + "ReportingEngine.SetPointColorDynamically Gold.docx"));
        }

        [Test]
        public void ConditionalExpressionForLeaveChartSeries()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Chart series.docx");
            
            int condition = 3;
            BuildReport(doc, new object[] { Common.GetManagers(), condition }, new[] { "managers", "condition" });

            doc.Save(ArtifactsDir + "ReportingEngine.TestLeaveChartSeries.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.TestLeaveChartSeries.docx", GoldsDir + "ReportingEngine.TestLeaveChartSeries Gold.docx"));
        }

        [Test, Ignore("WORDSNET-20810")]
        public void ConditionalExpressionForRemoveChartSeries()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Chart series.docx");
            
            int condition = 2;
            BuildReport(doc, new object[] { Common.GetManagers(), condition }, new[] { "managers", "condition" });

            doc.Save(ArtifactsDir + "ReportingEngine.TestRemoveChartSeries.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.TestRemoveChartSeries.docx", GoldsDir + "ReportingEngine.TestRemoveChartSeries Gold.docx"));
        }

        [Test]
        public void IndexOf()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Index of.docx");

            BuildReport(doc, Common.GetManagers(), "Managers");

            doc = DocumentHelper.SaveOpen(doc);

            Assert.AreEqual("The names are: John Smith, Tony Anderson, July James\f", doc.GetText());
        }

        [Test]
        public void IfElse()
        {
            Document doc = new Document(MyDir + "Reporting engine template - If-else.docx");

            BuildReport(doc, Common.GetManagers(), "m");

            doc = DocumentHelper.SaveOpen(doc);

            Assert.AreEqual("You have chosen 3 item(s).\f", doc.GetText());
        }

        [Test]
        public void IfElseWithoutData()
        {
            Document doc = new Document(MyDir + "Reporting engine template - If-else.docx");

            BuildReport(doc, Common.GetEmptyManagers(), "m");

            doc = DocumentHelper.SaveOpen(doc);

            Assert.AreEqual("You have chosen no items.\f", doc.GetText());
        }

        [Test]
        public void ExtensionMethods()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Extension methods.docx");

            BuildReport(doc, Common.GetManagers(), "Managers");

            doc.Save(ArtifactsDir + "ReportingEngine.ExtensionMethods.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.ExtensionMethods.docx", GoldsDir + "ReportingEngine.ExtensionMethods Gold.docx"));
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

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.Operators.docx", GoldsDir + "ReportingEngine.Operators Gold.docx"));
        }

        [Test]
        public void ContextualObjectMemberAccess()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Contextual object member access.docx");

            BuildReport(doc, Common.GetManagers(), "Managers");

            doc.Save(ArtifactsDir + "ReportingEngine.ContextualObjectMemberAccess.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.ContextualObjectMemberAccess.docx", GoldsDir + "ReportingEngine.ContextualObjectMemberAccess Gold.docx"));
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

            Assert.IsTrue(
                DocumentHelper.CompareDocs(
                    ArtifactsDir + "ReportingEngine.InsertDocumentDynamicallyWithAdditionalTemplateChecking.docx",
                    GoldsDir + "ReportingEngine.InsertDocumentDynamicallyWithAdditionalTemplateChecking Gold.docx"),
                "Fail inserting document by document");
        }

        [Test]
        public void InsertDocumentDynamicallyWithStyles()
        {
            Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.Document] -sourceStyles>>");

            DocumentTestClass doc = new DocumentTestBuilder()
                .WithDocument(new Document(MyDir + "Reporting engine template - Data table.docx")).Build();

            BuildReport(template, doc, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx", GoldsDir + "ReportingEngine.InsertDocumentDynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by document");
        }

        [Test]
        public void InsertDocumentDynamicallyByStream()
        {
            Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.DocumentStream]>>");

            DocumentTestClass docStream = new DocumentTestBuilder()
                .WithDocumentStream(new FileStream(mDocument, FileMode.Open, FileAccess.Read)).Build();

            BuildReport(template, docStream, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx", GoldsDir + "ReportingEngine.InsertDocumentDynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by stream");
        }

        [Test]
        public void InsertDocumentDynamicallyByBytes()
        {
            Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.DocumentBytes]>>");

            DocumentTestClass docBytes = new DocumentTestBuilder()
                .WithDocumentBytes(File.ReadAllBytes(MyDir + "Reporting engine template - Data table.docx")).Build();

            BuildReport(template, docBytes, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx", GoldsDir + "ReportingEngine.InsertDocumentDynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by bytes");
        }

        [Test]
        public void InsertDocumentDynamicallyByUri()
        {
            Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.DocumentString]>>");

            DocumentTestClass docUri = new DocumentTestBuilder()
                .WithDocumentString("http://www.snee.com/xml/xslt/sample.doc").Build();

            BuildReport(template, docUri, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx", GoldsDir + "ReportingEngine.InsertDocumentDynamically(uri) Gold.docx"), "Fail inserting document by uri");
        }

        [Test]
        public void InsertDocumentDynamicallyByBase64()
        {
            Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.DocumentString]>>");
            string base64Template = File.ReadAllText(MyDir + "Reporting engine template - Data table (base64).txt");

            DocumentTestClass docBase64 = new DocumentTestBuilder().WithDocumentString(base64Template).Build();

            BuildReport(template, docBase64, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.InsertDocumentDynamically.docx", GoldsDir + "ReportingEngine.InsertDocumentDynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by uri");
        }

        [Test]
        public void InsertImageDynamically()
        {
            Document template =
                DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.Image]>>", ShapeType.TextBox);
            
            #if NET462 || JAVA
            ImageTestClass image = new ImageTestBuilder().WithImage(Image.FromFile(mImage, true)).Build();
            #elif NETCOREAPP2_1 || __MOBILE__
            ImageTestClass image = new ImageTestBuilder().WithImage(SKBitmap.Decode(mImage)).Build();
            #endif
            
            BuildReport(template, image, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.InsertImageDynamically.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.InsertImageDynamically.docx", GoldsDir + "ReportingEngine.InsertImageDynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by bytes");
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

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.InsertImageDynamically.docx", GoldsDir + "ReportingEngine.InsertImageDynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by bytes");
        }

        [Test]
        public void InsertImageDynamicallyByBytes()
        {
            Document template =
                DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.ImageBytes]>>", ShapeType.TextBox);
            ImageTestClass imageBytes = new ImageTestBuilder().WithImageBytes(File.ReadAllBytes(mImage)).Build();

            BuildReport(template, imageBytes, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.InsertImageDynamically.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.InsertImageDynamically.docx", GoldsDir + "ReportingEngine.InsertImageDynamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by bytes");
        }

        [Test]
        public void InsertImageDynamicallyByUri()
        {
            Document template =
                DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.ImageString]>>", ShapeType.TextBox);
            ImageTestClass imageUri = new ImageTestBuilder()
                .WithImageString(
                    "http://joomla-aspose.dynabic.com/templates/aspose/App_Themes/V3/images/customers/americanexpress.png")
                .Build();

            BuildReport(template, imageUri, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.InsertImageDynamically.docx");

            Assert.IsTrue(
                DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.InsertImageDynamically.docx",
                    GoldsDir + "ReportingEngine.InsertImageDynamically(uri) Gold.docx"),
                "Fail inserting document by bytes");
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

            Assert.IsTrue(
                DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.InsertImageDynamically.docx",
                    GoldsDir + "ReportingEngine.InsertImageDynamically(stream,doc,bytes) Gold.docx"),
                "Fail inserting document by bytes");

        }
        
        [Test]
        public void DynamicStretchingImageWithinTextBox()
        {
            Document template = new Document(MyDir + "Reporting engine template - Dynamic stretching.docx");
            
#if NET462 || JAVA
            ImageTestClass image = new ImageTestBuilder().WithImage(Image.FromFile(mImage, true)).Build();
#elif NETCOREAPP2_1 || __MOBILE__
            ImageTestClass image = new ImageTestBuilder().WithImage(SKBitmap.Decode(mImage)).Build();
#endif
            BuildReport(template, image, "src", ReportBuildOptions.None);
            template.Save(ArtifactsDir + "ReportingEngine.DynamicStretchingImageWithinTextBox.docx");

            Assert.IsTrue(
                DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.DynamicStretchingImageWithinTextBox.docx",
                    GoldsDir + "ReportingEngine.DynamicStretchingImageWithinTextBox Gold.docx"));
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
            Assert.That(() => engine.BuildReport(doc, ""), Throws.TypeOf<InvalidOperationException>());
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

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.KnownTypes.docx", GoldsDir + "ReportingEngine.KnownTypes Gold.docx"));
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
                // Assert that the image is really insert in textbox 
                Assert.IsNotNull(shape.Fill.ImageBytes);

                // Assert that the width is preserved, and the height is changed
                Assert.AreNotEqual(346.35, shape.Height);
                Assert.AreEqual(431.5, shape.Width);
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
                Assert.IsNotNull(shape.Fill.ImageBytes);

                // Assert that the height is preserved, and the width is changed
                Assert.AreNotEqual(431.5, shape.Width);
                Assert.AreEqual(346.35, shape.Height);
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
                Assert.IsNotNull(shape.Fill.ImageBytes);

                // Assert that the height and the width are changed
                Assert.AreNotEqual(346.35, shape.Height);
                Assert.AreNotEqual(431.5, shape.Width);
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
                Assert.IsNotNull(shape.Fill.ImageBytes);

                // Assert that textbox size are equal image size
                Assert.AreEqual(300.0d, shape.Height);
                Assert.AreEqual(300.0d, shape.Width);
            }
        }

        [Test]
        public void WithoutMissingMembers()
        {
            DocumentBuilder builder = new DocumentBuilder();

            //Add templete to the document for reporting engine
            DocumentHelper.InsertBuilderText(builder,
                new[] { "<<[missingObject.First().id]>>", "<<foreach [in missingObject]>><<[id]>><</foreach>>" });

            //Assert that build report failed without "ReportBuildOptions.AllowMissingMembers"
            Assert.That(() => BuildReport(builder.Document, new DataSet(), "", ReportBuildOptions.None),
                Throws.TypeOf<InvalidOperationException>());
        }

        [Test]
        public void WithMissingMembers()
        {
            DocumentBuilder builder = new DocumentBuilder();

            //Add templete to the document for reporting engine
            DocumentHelper.InsertBuilderText(builder,
                new[] { "<<[missingObject.First().id]>>", "<<foreach [in missingObject]>><<[id]>><</foreach>>" });

            BuildReport(builder.Document, new DataSet(), "", ReportBuildOptions.AllowMissingMembers);

            //Assert that build report success with "ReportBuildOptions.AllowMissingMembers"
            Assert.AreEqual(ControlChar.ParagraphBreak + ControlChar.ParagraphBreak + ControlChar.SectionBreak,
                builder.Document.GetText());
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
        public void SetBackgroundColor()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Background color.docx");

            List<ColorItemTestClass> colors = new List<ColorItemTestClass>
            {
                new ColorItemTestBuilder().WithColor("Black", Color.Black).Build(),
                new ColorItemTestBuilder().WithColor("Red", Color.FromArgb(255, 0, 0)).Build(),
                new ColorItemTestBuilder().WithColor("Empty", Color.Empty).Build()
            };

            BuildReport(doc, colors, "Colors");

            doc.Save(ArtifactsDir + "ReportingEngine.BackColor.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.BackColor.docx",
                GoldsDir + "ReportingEngine.BackColor Gold.docx"));
        }

        [Test]
        public void DoNotRemoveEmptyParagraphs()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Remove empty paragraphs.docx");

            BuildReport(doc, Common.GetManagers(), "Managers");

            doc.Save(ArtifactsDir + "ReportingEngine.DoNotRemoveEmptyParagraphs.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.DoNotRemoveEmptyParagraphs.docx",
                GoldsDir + "ReportingEngine.DoNotRemoveEmptyParagraphs Gold.docx"));
        }

        [Test]
        public void RemoveEmptyParagraphs()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Remove empty paragraphs.docx");

            BuildReport(doc, Common.GetManagers(), "Managers", ReportBuildOptions.RemoveEmptyParagraphs);

            doc.Save(ArtifactsDir + "ReportingEngine.RemoveEmptyParagraphs.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.RemoveEmptyParagraphs.docx",
                GoldsDir + "ReportingEngine.RemoveEmptyParagraphs Gold.docx"));
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

            Assert.IsTrue(DocumentHelper.CompareDocs(artifactPath, goldPath));
        }

        [Test]
        public void XmlDataStringWithoutSchema()
        {
            Document doc = new Document(MyDir + "Reporting engine template - XML data destination.docx");

            XmlDataSource dataSource = new XmlDataSource(MyDir + "List of people.xml");
            BuildReport(doc, dataSource, "persons");

            doc.Save(ArtifactsDir + "ReportingEngine.XmlDataString.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.XmlDataString.docx",
                GoldsDir + "ReportingEngine.DataSource Gold.docx"));
        }

        [Test]
        public void XmlDataStreamWithoutSchema()
        {
            Document doc = new Document(MyDir + "Reporting engine template - XML data destination.docx");

            using (FileStream stream = File.OpenRead(MyDir + "List of people.xml"))
            {
                XmlDataSource dataSource = new XmlDataSource(stream);
                BuildReport(doc, dataSource, "persons");
            }

            doc.Save(ArtifactsDir + "ReportingEngine.XmlDataStream.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.XmlDataStream.docx",
                GoldsDir + "ReportingEngine.DataSource Gold.docx"));
        }

        [Test]
        public void XmlDataWithNestedElements()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Data destination with nested elements.docx");

            XmlDataSource dataSource = new XmlDataSource(MyDir + "Nested elements.xml");
            BuildReport(doc, dataSource, "managers");

            doc.Save(ArtifactsDir + "ReportingEngine.XmlDataWithNestedElements.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.XmlDataWithNestedElements.docx",
                GoldsDir + "ReportingEngine.DataSourceWithNestedElements Gold.docx"));
        }

        [Test]
        public void JsonDataString()
        {
            Document doc = new Document(MyDir + "Reporting engine template - JSON data destination.docx");

            JsonDataLoadOptions options = new JsonDataLoadOptions();
            options.ExactDateTimeParseFormat = "MM/dd/yyyy";

            JsonDataSource dataSource = new JsonDataSource(MyDir + "List of people.json", options);
            BuildReport(doc, dataSource, "persons");
            
            doc.Save(ArtifactsDir + "ReportingEngine.JsonDataString.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.JsonDataString.docx",
                GoldsDir + "ReportingEngine.JsonDataString Gold.docx"));
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
            Document doc = new Document(MyDir + "Reporting engine template - JSON data destination.docx");
            
            JsonDataLoadOptions options = new JsonDataLoadOptions();
            options.ExactDateTimeParseFormat = "MM/dd/yyyy";
            
            using (FileStream stream = File.OpenRead(MyDir + "List of people.json"))
            {
                JsonDataSource dataSource = new JsonDataSource(stream, options);
                BuildReport(doc, dataSource, "persons");
            }

            doc.Save(ArtifactsDir + "ReportingEngine.JsonDataStream.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.JsonDataStream.docx",
                GoldsDir + "ReportingEngine.JsonDataString Gold.docx"));
        }

        [Test]
        public void JsonDataWithNestedElements()
        {
            Document doc = new Document(MyDir + "Reporting engine template - Data destination with nested elements.docx");

            JsonDataSource dataSource = new JsonDataSource(MyDir + "Nested elements.json");
            BuildReport(doc, dataSource, "managers");
            
            doc.Save(ArtifactsDir + "ReportingEngine.JsonDataWithNestedElements.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.JsonDataWithNestedElements.docx",
                GoldsDir + "ReportingEngine.DataSourceWithNestedElements Gold.docx"));
        }

        [Test]
        public void CsvDataString()
        {
            Document doc = new Document(MyDir + "Reporting engine template - CSV data destination.docx");
            
            CsvDataLoadOptions loadOptions = new CsvDataLoadOptions(true);
            loadOptions.Delimiter = ';';
            loadOptions.CommentChar = '$';

            CsvDataSource dataSource = new CsvDataSource(MyDir + "List of people.csv", loadOptions);
            BuildReport(doc, dataSource, "persons");
            
            doc.Save(ArtifactsDir + "ReportingEngine.CsvDataString.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.CsvDataString.docx",
                GoldsDir + "ReportingEngine.CsvData Gold.docx"));
        }

        [Test]
        public void CsvDataStream()
        {
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

            Assert.IsTrue(DocumentHelper.CompareDocs(ArtifactsDir + "ReportingEngine.CsvDataStream.docx",
                GoldsDir + "ReportingEngine.CsvData Gold.docx"));
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

            Assert.AreEqual(expectedItems.Length, sdt.ListItems.Count);

            for (int i = 0; i < expectedItems.Length; i++)
            {
                Assert.AreEqual(expectedItems[i].Value, sdt.ListItems[i].Value);
                Assert.AreEqual(expectedItems[i].DisplayText, sdt.ListItems[i].DisplayText);
            }
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