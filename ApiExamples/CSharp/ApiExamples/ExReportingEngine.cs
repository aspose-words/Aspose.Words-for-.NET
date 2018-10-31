﻿// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
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
using Aspose.Words.Reporting;
using NUnit.Framework;

#if NETSTANDARD2_0 || __MOBILE__
using SkiaSharp;
#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExReportingEngine : ApiExampleBase
    {
        private readonly string mImage = ImageDir + "Test_636_852.gif";
        private readonly string mDocument = MyDir + "ReportingEngine.TestDataTable.docx";

        [Test]
        public void SimpleCase()
        {
            Document doc = DocumentHelper.CreateSimpleDocument("<<[s.Name]>> says: <<[s.Message]>>");

            MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "Hello World");
            BuildReport(doc, sender, "s", ReportBuildOptions.None);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual("LINQ Reporting Engine says: Hello World\f", doc.GetText());
        }

        [Test]
        public void StringFormat()
        {
            Document doc = DocumentHelper.CreateSimpleDocument(
                "<<[s.Name]:lower>> says: <<[s.Message]:upper>>, <<[s.Message]:caps>>, <<[s.Message]:firstCap>>");

            MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "hello world");
            BuildReport(doc, sender, "s");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

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

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual("A : ii, 200th, FIRST, Two, C8, - 200 -\f", doc.GetText());
        }

        [Test]
        public void DataTableTest()
        {
            Document doc = new Document(MyDir + "ReportingEngine.TestDataTable.docx");

            BuildReport(doc, Common.GetContracts(), "Contracts");
            doc.Save(MyDir + @"\Artifacts\ReportingEngine.TestDataTable.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.TestDataTable.docx",
                MyDir + @"\Golds\ReportingEngine.TestDataTable Gold.docx"));
        }

        [Test]
        public void ProgressiveTotal()
        {
            Document doc = new Document(MyDir + "ReportingEngine.Total.docx");

            BuildReport(doc, Common.GetContracts(), "Contracts");

            doc.Save(MyDir + @"\Artifacts\ReportingEngine.Total.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.Total.docx",
                MyDir + @"\Golds\ReportingEngine.Total Gold.docx"));
        }

        [Test]
        public void NestedDataTableTest()
        {
            Document doc = new Document(MyDir + "ReportingEngine.TestNestedDataTable.docx");

            BuildReport(doc, Common.GetManagers(), "Managers");

            doc.Save(MyDir + @"\Artifacts\ReportingEngine.TestNestedDataTable.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.TestNestedDataTable.docx",
                MyDir + @"\Golds\ReportingEngine.TestNestedDataTable Gold.docx"));
        }

        [Test]
        public void ChartTest()
        {
            Document doc = new Document(MyDir + "ReportingEngine.TestChart.docx");

            BuildReport(doc, Common.GetManagers(), "managers");

            doc.Save(MyDir + @"\Artifacts\ReportingEngine.TestChart.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.TestChart.docx",
                MyDir + @"\Golds\ReportingEngine.TestChart Gold.docx"));
        }

        [Test]
        public void BubbleChartTest()
        {
            Document doc = new Document(MyDir + "ReportingEngine.TestBubbleChart.docx");

            BuildReport(doc, Common.GetManagers(), "managers");

            doc.Save(MyDir + @"\Artifacts\ReportingEngine.TestBubbleChart.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.TestBubbleChart.docx",
                MyDir + @"\Golds\ReportingEngine.TestBubbleChart Gold.docx"));
        }

        [Test]
        public void SetChartSeriesColorsDynamically()
        {
            Document doc = new Document(MyDir + "ReportingEngine.SetChartSeriesColorDinamically.docx");

            BuildReport(doc, Common.GetManagers(), "managers");

            doc.Save(MyDir + @"\Artifacts\ReportingEngine.SetChartSeriesColorDinamically.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(
                MyDir + @"\Artifacts\ReportingEngine.SetChartSeriesColorDinamically.docx",
                MyDir + @"\Golds\ReportingEngine.SetChartSeriesColorDinamically Gold.docx"));
        }

        [Test]
        public void SetPointColorsDynamically()
        {
            Document doc = new Document(MyDir + "ReportingEngine.SetPointColorDinamically.docx");

            List<ColorItemTestClass> colors = new List<ColorItemTestClass>
            {
                new ColorItemTestBuilder().WithColorCodeAndValues("Black", Color.Black.ToArgb(), 1.0, 2.5, 3.5).Build(),
                new ColorItemTestBuilder().WithColorCodeAndValues("Red", Color.Red.ToArgb(), 2.0, 4.0, 2.5).Build(),
                new ColorItemTestBuilder().WithColorCodeAndValues("Green", Color.Green.ToArgb(), 0.5, 1.5, 2.5).Build(),
                new ColorItemTestBuilder().WithColorCodeAndValues("Blue", Color.Blue.ToArgb(), 4.5, 3.5, 1.5).Build(),
                new ColorItemTestBuilder().WithColorCodeAndValues("Yellow", Color.Yellow.ToArgb(), 5.0, 2.5, 1.5)
                    .Build()
            };

            BuildReport(doc, colors, "colorItems", new[] { typeof(ColorItemTestClass) });
            doc.Save(MyDir + @"\Artifacts\ReportingEngine.SetPointColorDinamically.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(
                MyDir + @"\Artifacts\ReportingEngine.SetPointColorDinamically.docx",
                MyDir + @"\Golds\ReportingEngine.SetPointColorDinamically Gold.docx"));
        }

        [Test]
        public void ConditionalExpressionForLeaveChartSeries()
        {
            int condition = 3;

            Document doc = new Document(MyDir + "ReportingEngine.TestRemoveChartSeries.docx");

            BuildReport(doc, new object[] { Common.GetManagers(), condition }, new[] { "managers", "condition" });
            doc.Save(MyDir + @"\Artifacts\ReportingEngine.TestLeaveChartSeries.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.TestLeaveChartSeries.docx",
                MyDir + @"\Golds\ReportingEngine.TestLeaveChartSeries Gold.docx"));
        }

        [Test]
        public void ConditionalExpressionForRemoveChartSeries()
        {
            int condition = 2;

            Document doc = new Document(MyDir + "ReportingEngine.TestRemoveChartSeries.docx");

            BuildReport(doc, new object[] { Common.GetManagers(), condition }, new[] { "managers", "condition" });
            doc.Save(MyDir + @"\Artifacts\ReportingEngine.TestRemoveChartSeries.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.TestRemoveChartSeries.docx",
                MyDir + @"\Golds\ReportingEngine.TestRemoveChartSeries Gold.docx"));
        }

        [Test]
        public void IndexOf()
        {
            Document doc = new Document(MyDir + "ReportingEngine.TestIndexOf.docx");

            BuildReport(doc, Common.GetManagers(), "Managers");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual("The names are: John Smith, Tony Anderson, July James\f", doc.GetText());
        }

        [Test]
        public void IfElse()
        {
            Document doc = new Document(MyDir + "ReportingEngine.IfElse.docx");

            BuildReport(doc, Common.GetManagers(), "m");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual("You have chosen 3 item(s).\f", doc.GetText());
        }

        [Test]
        public void IfElseWithoutData()
        {
            Document doc = new Document(MyDir + "ReportingEngine.IfElse.docx");

            BuildReport(doc, Common.GetEmptyManagers(), "m");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual("You have chosen no items.\f", doc.GetText());
        }

        [Test]
        public void ExtensionMethods()
        {
            Document doc = new Document(MyDir + "ReportingEngine.ExtensionMethods.docx");

            BuildReport(doc, Common.GetManagers(), "Managers");

            doc.Save(MyDir + @"\Artifacts\ReportingEngine.ExtensionMethods.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.ExtensionMethods.docx",
                MyDir + @"\Golds\ReportingEngine.ExtensionMethods Gold.docx"));
        }

        [Test]
        public void Operators()
        {
            Document doc = new Document(MyDir + "ReportingEngine.Operators.docx");

            NumericTestClass testData = new NumericTestBuilder().WithValuesAndLogical(1, 2.0, 3, null, true).Build();

            ReportingEngine report = new ReportingEngine();
            report.KnownTypes.Add(typeof(NumericTestBuilder));
            report.BuildReport(doc, testData, "ds");

            doc.Save(MyDir + @"\Artifacts\ReportingEngine.Operators.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.Operators.docx",
                MyDir + @"\Golds\ReportingEngine.Operators Gold.docx"));
        }

        [Test]
        public void ContextualObjectMemberAccess()
        {
            Document doc = new Document(MyDir + "ReportingEngine.ContextualObjectMemberAccess.docx");

            BuildReport(doc, Common.GetManagers(), "Managers");

            doc.Save(MyDir + @"\Artifacts\ReportingEngine.ContextualObjectMemberAccess.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(
                MyDir + @"\Artifacts\ReportingEngine.ContextualObjectMemberAccess.docx",
                MyDir + @"\Golds\ReportingEngine.ContextualObjectMemberAccess Gold.docx"));
        }

        [Test]
        public void InsertDocumentDinamicallyWithAdditionalTemplateChecking()
        {
            Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.Document] -build>>");

            DocumentTestClass doc = new DocumentTestBuilder()
                .WithDocument(new Document(MyDir + "ReportingEngine.TestDataTable.docx")).Build();

            BuildReport(template, new object[] { doc, Common.GetContracts() }, new[] { "src", "Contracts" }, 
                ReportBuildOptions.None);
            template.Save(
                MyDir + @"\Artifacts\ReportingEngine.InsertDocumentDinamicallyWithAdditionalTemplateChecking.docx");

            Assert.IsTrue(
                DocumentHelper.CompareDocs(
                    MyDir + @"\Artifacts\ReportingEngine.InsertDocumentDinamicallyWithAdditionalTemplateChecking.docx",
                    MyDir + @"\Golds\ReportingEngine.InsertDocumentDinamicallyWithAdditionalTemplateChecking Gold.docx"),
                "Fail inserting document by document");
        }


        [Test]
        public void InsertDocumentDinamically()
        {
            Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.Document]>>");

            DocumentTestClass doc = new DocumentTestBuilder()
                .WithDocument(new Document(MyDir + "ReportingEngine.TestDataTable.docx")).Build();

            BuildReport(template, doc, "src", ReportBuildOptions.None);
            template.Save(MyDir + @"\Artifacts\ReportingEngine.InsertDocumentDinamically.docx");

            Assert.IsTrue(
                DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.InsertDocumentDinamically.docx",
                    MyDir + @"\Golds\ReportingEngine.InsertDocumentDinamically(stream,doc,bytes) Gold.docx"),
                "Fail inserting document by document");
        }

        [Test]
        public void InsertDocumentDinamicallyByStream()
        {
            Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.DocumentStream]>>");

            DocumentTestClass docStream = new DocumentTestBuilder()
                .WithDocumentStream(new FileStream(mDocument, FileMode.Open, FileAccess.Read)).Build();

            BuildReport(template, docStream, "src", ReportBuildOptions.None);
            template.Save(MyDir + @"\Artifacts\ReportingEngine.InsertDocumentDinamically.docx");

            Assert.IsTrue(
                DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.InsertDocumentDinamically.docx",
                    MyDir + @"\Golds\ReportingEngine.InsertDocumentDinamically(stream,doc,bytes) Gold.docx"),
                "Fail inserting document by stream");
        }

        [Test]
        public void InsertDocumentDinamicallyByBytes()
        {
            Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.DocumentBytes]>>");

            DocumentTestClass docBytes = new DocumentTestBuilder()
                .WithDocumentBytes(File.ReadAllBytes(MyDir + "ReportingEngine.TestDataTable.docx")).Build();

            BuildReport(template, docBytes, "src", ReportBuildOptions.None);
            template.Save(MyDir + @"\Artifacts\ReportingEngine.InsertDocumentDinamically.docx");

            Assert.IsTrue(
                DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.InsertDocumentDinamically.docx",
                    MyDir + @"\Golds\ReportingEngine.InsertDocumentDinamically(stream,doc,bytes) Gold.docx"),
                "Fail inserting document by bytes");
        }

        [Test]
        public void InsertDocumentDinamicallyByUri()
        {
            Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.DocumentUri]>>");

            DocumentTestClass docUri = new DocumentTestBuilder()
                .WithDocumentUri("http://www.snee.com/xml/xslt/sample.doc").Build();

            BuildReport(template, docUri, "src", ReportBuildOptions.None);
            template.Save(MyDir + @"\Artifacts\ReportingEngine.InsertDocumentDinamically.docx");

            Assert.IsTrue(
                DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.InsertDocumentDinamically.docx",
                    MyDir + @"\Golds\ReportingEngine.InsertDocumentDinamically(uri) Gold.docx"),
                "Fail inserting document by uri");
        }

        [Test]
        public void InsertImageDinamically()
        {
            Document template =
                DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.Image]>>", ShapeType.TextBox);
#if NETSTANDARD2_0 || __MOBILE__
            ImageTestClass image = new ImageTestBuilder().WithImage(SKBitmap.Decode(mImage)).Build();
#else
            ImageTestClass image = new ImageTestBuilder().WithImage(Image.FromFile(mImage, true)).Build();
#endif

            BuildReport(template, image, "src", ReportBuildOptions.None);
            template.Save(MyDir + @"\Artifacts\ReportingEngine.InsertImageDinamically.docx");

            Assert.IsTrue(
                DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.InsertImageDinamically.docx",
                    MyDir + @"\Golds\ReportingEngine.InsertImageDinamically(stream,doc,bytes) Gold.docx"),
                "Fail inserting document by bytes");
        }

        [Test]
        public void InsertImageDinamicallyByStream()
        {
            Document template =
                DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.ImageStream]>>", ShapeType.TextBox);
            ImageTestClass imageStream = new ImageTestBuilder()
                .WithImageStream(new FileStream(mImage, FileMode.Open, FileAccess.Read)).Build();

            BuildReport(template, imageStream, "src", ReportBuildOptions.None);
            template.Save(MyDir + @"\Artifacts\ReportingEngine.InsertImageDinamically.docx");

            Assert.IsTrue(
                DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.InsertImageDinamically.docx",
                    MyDir + @"\Golds\ReportingEngine.InsertImageDinamically(stream,doc,bytes) Gold.docx"),
                "Fail inserting document by bytes");
        }

        [Test]
        public void InsertImageDinamicallyByBytes()
        {
            Document template =
                DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.ImageBytes]>>", ShapeType.TextBox);
            ImageTestClass imageBytes = new ImageTestBuilder().WithImageBytes(File.ReadAllBytes(mImage)).Build();

            BuildReport(template, imageBytes, "src", ReportBuildOptions.None);
            template.Save(MyDir + @"\Artifacts\ReportingEngine.InsertImageDinamically.docx");

            Assert.IsTrue(
                DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.InsertImageDinamically.docx",
                    MyDir + @"\Golds\ReportingEngine.InsertImageDinamically(stream,doc,bytes) Gold.docx"),
                "Fail inserting document by bytes");
        }

        [Test]
        public void InsertImageDinamicallyByUri()
        {
            Document template =
                DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.ImageUri]>>", ShapeType.TextBox);
            ImageTestClass imageUri = new ImageTestBuilder()
                .WithImageUri(
                    "http://joomla-aspose.dynabic.com/templates/aspose/App_Themes/V3/images/customers/americanexpress.png")
                .Build();

            BuildReport(template, imageUri, "src", ReportBuildOptions.None);
            template.Save(MyDir + @"\Artifacts\ReportingEngine.InsertImageDinamically.docx");

            Assert.IsTrue(
                DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.InsertImageDinamically.docx",
                    MyDir + @"\Golds\ReportingEngine.InsertImageDinamically(uri) Gold.docx"),
                "Fail inserting document by bytes");
        }

        [Test]
        public void InsertHyperlinksDinamically()
        {
            Document template = new Document(MyDir + "ReportingEngine.InsertingHyperlinks.docx");
            BuildReport(template, 
                new Object[]
                {
                    "https://auckland.dynabic.com/wiki/display/org/Supported+dynamic+insertion+of+hyperlinks+for+LINQ+Reporting+Engine",
                    "Aspose"
                },
                new[]
                {
                    "uri_expression", 
                    "display_text_expression"
                });

            template.Save(MyDir + @"\Artifacts\ReportingEngine.InsertHyperlinksDinamically.docx");

            Assert.IsTrue(
                DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.InsertHyperlinksDinamically.docx",
                    MyDir + @"\Golds\ReportingEngine.InsertHyperlinksDinamically Gold.docx"),
                "Fail inserting document by bytes");
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

            BuildReport(doc, "", new[] { typeof(DateTime) });
            doc.Save(MyDir + @"\Artifacts\ReportingEngine.KnownTypes.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.KnownTypes.docx",
                MyDir + @"\Golds\ReportingEngine.KnownTypes Gold.docx"));
        }

        [Test]
        public void WorkWithSingleColumnTableRow()
        {
            Document doc = new Document(MyDir + "ReportingEngine.SingleColumnTableRow.docx");
            BuildReport(doc, Common.GetManagers(), "Managers");

            doc.Save(MyDir + @"\Artifacts\ReportingEngine.SingleColumnTableRow.docx");
        }

        [Test]
        public void WorkWithSingleColumnTableRowGreedy()
        {
            Document doc = new Document(MyDir + "ReportingEngine.SingleColumnTableRowGreedy.docx");
            BuildReport(doc, Common.GetManagers(), "Managers");

            doc.Save(MyDir + @"\Artifacts\ReportingEngine.SingleColumnTableRowGreedy.docx");
        }

        [Test]
        public void TableRowConditionalBlocks()
        {
            Document doc = new Document(MyDir + "TableRowConditionalBlocks.docx");

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

            doc.Save(MyDir + @"\Artifacts\ReportingEngine.TableRowConditionalBlocks.docx");
        }

        [Test]
        public void IfGreedy()
        {
            Document doc = new Document(MyDir + "ReportingEngine.IfGreedy.docx");

            AsposeData obj = new AsposeData
            {
                List = new List<string>
                {
                    "abc"
                }
            };

            BuildReport(doc, obj);

            doc.Save(MyDir + @"\Artifacts\IfGreedy.docx");
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

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            doc = new Document(dstStream);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes.OfType<Shape>())
            {
                // Assert that the image is really insert in textbox 
                Assert.IsNotNull(shape.Fill.ImageBytes);

                // Assert that width is keeped and height is changed
                Assert.AreNotEqual(346.35, shape.Height);
                Assert.AreEqual(431.5, shape.Width);
            }

            dstStream.Dispose();
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

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            doc = new Document(dstStream);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes.OfType<Shape>())
            {
                // Assert that the image is really insert in textbox and 
                Assert.IsNotNull(shape.Fill.ImageBytes);

                // Assert that height is keeped and width is changed
                Assert.AreNotEqual(431.5, shape.Width);
                Assert.AreEqual(346.35, shape.Height);
            }

            dstStream.Dispose();
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

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            doc = new Document(dstStream);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes.OfType<Shape>())
            {
                // Assert that the image is really insert in textbox 
                Assert.IsNotNull(shape.Fill.ImageBytes);

                // Assert that height is changed and width is changed
                Assert.AreNotEqual(346.35, shape.Height);
                Assert.AreNotEqual(431.5, shape.Width);
            }

            dstStream.Dispose();
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

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            doc = new Document(dstStream);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes.OfType<Shape>())
            {
                // Assert that the image is really insert in textbox 
                Assert.IsNotNull(shape.Fill.ImageBytes);

                // Assert that textbox size are equal image size
                Assert.AreEqual(346.35, shape.Height);
                Assert.AreEqual(258.54, shape.Width);
            }

            dstStream.Dispose();
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

        [Test]
        public void SetBackgroundColor()
        {
            Document doc = new Document(MyDir + "ReportingEngine.BackColor.docx");

            List<ColorItemTestClass> colors = new List<ColorItemTestClass>
            {
                new ColorItemTestBuilder().WithColor("Black", Color.Black).Build(),
                new ColorItemTestBuilder().WithColor("Red", Color.FromArgb(255, 0, 0)).Build(),
                new ColorItemTestBuilder().WithColor("Empty", Color.Empty).Build()
            };

            BuildReport(doc, colors, "Colors");

            doc.Save(MyDir + @"\Artifacts\ReportingEngine.BackColor.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.BackColor.docx",
                MyDir + @"\Golds\ReportingEngine.BackColor Gold.docx"));
        }

        [Test]
        public void DoNotRemoveEmptyParagraphs()
        {
            Document doc = new Document(MyDir + "ReportingEngine.RemoveEmptyParagraphs.docx");

            BuildReport(doc, Common.GetManagers(), "Managers");

            doc.Save(MyDir + @"\Artifacts\ReportingEngine.DoNotRemoveEmptyParagraphs.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.DoNotRemoveEmptyParagraphs.docx",
                MyDir + @"\Golds\ReportingEngine.DoNotRemoveEmptyParagraphs Gold.docx"));
        }

        [Test]
        public void RemoveEmptyParagraphs()
        {
            Document doc = new Document(MyDir + "ReportingEngine.RemoveEmptyParagraphs.docx");

            BuildReport(doc, Common.GetManagers(), "Managers", ReportBuildOptions.RemoveEmptyParagraphs);

            doc.Save(MyDir + @"\Artifacts\ReportingEngine.RemoveEmptyParagraphs.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.RemoveEmptyParagraphs.docx",
                MyDir + @"\Golds\ReportingEngine.RemoveEmptyParagraphs Gold.docx"));
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