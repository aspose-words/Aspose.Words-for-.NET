// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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
using System.Text.RegularExpressions;
using Aspose.Page.XPS;
using Aspose.Page.XPS.XpsModel;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.LowCode;
using Aspose.Words.Replacing;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using NUnit.Framework;
using LoadOptions = Aspose.Words.Loading.LoadOptions;

namespace ApiExamples
{
    [TestFixture]
    public class ExLowCode : ApiExampleBase
    {
        [Test]
        public void MergeDocuments()
        {
            //ExStart
            //ExFor:Merger.Merge(String, String[])
            //ExFor:Merger.Merge(String[], MergeFormatMode)
            //ExFor:Merger.Merge(String[], LoadOptions[], MergeFormatMode)
            //ExFor:Merger.Merge(String, String[], SaveOptions, MergeFormatMode)
            //ExFor:Merger.Merge(String, String[], SaveFormat, MergeFormatMode)
            //ExFor:Merger.Merge(String, String[], LoadOptions[], SaveOptions, MergeFormatMode)
            //ExFor:LowCode.MergeFormatMode
            //ExFor:LowCode.Merger
            //ExSummary:Shows how to merge documents into a single output document.
            //There is a several ways to merge documents:
            string inputDoc1 = MyDir + "Big document.docx";
            string inputDoc2 = MyDir + "Tables.docx";

            Merger.Merge(ArtifactsDir + "LowCode.MergeDocument.1.docx", new[] { inputDoc1, inputDoc2 });

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.Password = "Aspose.Words";
            Merger.Merge(ArtifactsDir + "LowCode.MergeDocument.2.docx", new[] { inputDoc1, inputDoc2 }, saveOptions, MergeFormatMode.KeepSourceFormatting);

            Merger.Merge(ArtifactsDir + "LowCode.MergeDocument.3.pdf", new[] { inputDoc1, inputDoc2 }, SaveFormat.Pdf, MergeFormatMode.KeepSourceLayout);

            LoadOptions firstLoadOptions = new LoadOptions();
            firstLoadOptions.IgnoreOleData = true;
            LoadOptions secondLoadOptions = new LoadOptions();
            secondLoadOptions.IgnoreOleData = false;
            Merger.Merge(ArtifactsDir + "LowCode.MergeDocument.4.docx", new[] { inputDoc1, inputDoc2 }, new[] { firstLoadOptions, secondLoadOptions }, saveOptions, MergeFormatMode.KeepSourceFormatting);

            Document doc = Merger.Merge(new[] { inputDoc1, inputDoc2 }, MergeFormatMode.MergeFormatting);
            doc.Save(ArtifactsDir + "LowCode.MergeDocument.5.docx");

            doc = Merger.Merge(new[] { inputDoc1, inputDoc2 }, new[] { firstLoadOptions, secondLoadOptions }, MergeFormatMode.MergeFormatting);
            doc.Save(ArtifactsDir + "LowCode.MergeDocument.6.docx");
            //ExEnd
        }

        [Test]
        public void MergeContextDocuments()
        {
            //ExStart:MergeContextDocuments
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Processor
            //ExFor:Processor.From(String, LoadOptions)
            //ExFor:Processor.To(String, SaveOptions)
            //ExFor:Processor.To(String, SaveFormat)
            //ExFor:Processor.Execute
            //ExFor:Merger.Create(MergerContext)
            //ExFor:MergerContext
            //ExSummary:Shows how to merge documents into a single output document using context.
            //There is a several ways to merge documents:
            string inputDoc1 = MyDir + "Big document.docx";
            string inputDoc2 = MyDir + "Tables.docx";
            MergerContext context = new MergerContext();
            context.MergeFormatMode = MergeFormatMode.KeepSourceFormatting;

            Merger.Create(context)
                .From(inputDoc1)
                .From(inputDoc2)
                .To(ArtifactsDir + "LowCode.MergeContextDocuments.1.docx")
                .Execute();

            LoadOptions firstLoadOptions = new LoadOptions();
            firstLoadOptions.IgnoreOleData = true;
            LoadOptions secondLoadOptions = new LoadOptions();
            secondLoadOptions.IgnoreOleData = false;
            MergerContext context2 = new MergerContext();
            context2.MergeFormatMode = MergeFormatMode.KeepSourceFormatting;
            Merger.Create(context2)
                .From(inputDoc1, firstLoadOptions)
                .From(inputDoc2, secondLoadOptions)
                .To(ArtifactsDir + "LowCode.MergeContextDocuments.2.docx", SaveFormat.Docx)
                .Execute();

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.Password = "Aspose.Words";
            MergerContext context3 = new MergerContext();
            context3.MergeFormatMode = MergeFormatMode.KeepSourceFormatting;
            Merger.Create(context3)
                .From(inputDoc1)
                .From(inputDoc2)
                .To(ArtifactsDir + "LowCode.MergeContextDocuments.3.docx", saveOptions)
                .Execute();
            //ExEnd:MergeContextDocuments
        }

        [Test]
        public void MergeStreamDocument()
        {
            //ExStart
            //ExFor:Merger.Merge(Stream[], MergeFormatMode)
            //ExFor:Merger.Merge(Stream[], LoadOptions[], MergeFormatMode)
            //ExFor:Merger.Merge(Stream, Stream[], SaveOptions, MergeFormatMode)
            //ExFor:Merger.Merge(Stream, Stream[], LoadOptions[], SaveOptions, MergeFormatMode)
            //ExFor:Merger.Merge(Stream, Stream[], SaveFormat)
            //ExSummary:Shows how to merge documents from stream into a single output document.
            //There is a several ways to merge documents from stream:
            using (FileStream firstStreamIn = new FileStream(MyDir + "Big document.docx", FileMode.Open, FileAccess.Read))
            {
                using (FileStream secondStreamIn = new FileStream(MyDir + "Tables.docx", FileMode.Open, FileAccess.Read))
                {
                    OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
                    saveOptions.Password = "Aspose.Words";
                    using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MergeStreamDocument.1.docx", FileMode.Create, FileAccess.ReadWrite))
                        Merger.Merge(streamOut, new[] { firstStreamIn, secondStreamIn }, saveOptions, MergeFormatMode.KeepSourceFormatting);

                    using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MergeStreamDocument.2.docx", FileMode.Create, FileAccess.ReadWrite))
                        Merger.Merge(streamOut, new[] { firstStreamIn, secondStreamIn }, SaveFormat.Docx);

                    LoadOptions firstLoadOptions = new LoadOptions();
                    firstLoadOptions.IgnoreOleData = true;
                    LoadOptions secondLoadOptions = new LoadOptions();
                    secondLoadOptions.IgnoreOleData = false;
                    using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MergeStreamDocument.3.docx", FileMode.Create, FileAccess.ReadWrite))
                        Merger.Merge(streamOut, new[] { firstStreamIn, secondStreamIn }, new[] { firstLoadOptions, secondLoadOptions }, saveOptions, MergeFormatMode.KeepSourceFormatting);

                    Document firstDoc = Merger.Merge(new[] { firstStreamIn, secondStreamIn }, MergeFormatMode.MergeFormatting);
                    firstDoc.Save(ArtifactsDir + "LowCode.MergeStreamDocument.4.docx");

                    Document secondDoc = Merger.Merge(new[] { firstStreamIn, secondStreamIn }, new[] { firstLoadOptions, secondLoadOptions }, MergeFormatMode.MergeFormatting);
                    secondDoc.Save(ArtifactsDir + "LowCode.MergeStreamDocument.5.docx");
                }
            }
            //ExEnd
        }

        [Test]
        public void MergeStreamContextDocuments()
        {
            //ExStart:MergeStreamContextDocuments
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Processor
            //ExFor:Processor.From(Stream, LoadOptions)
            //ExFor:Processor.To(Stream, SaveFormat)
            //ExFor:Processor.To(Stream, SaveOptions)
            //ExFor:Processor.Execute
            //ExFor:Merger.Create(MergerContext)
            //ExFor:MergerContext
            //ExSummary:Shows how to merge documents from stream into a single output document using context.
            //There is a several ways to merge documents:
            string inputDoc1 = MyDir + "Big document.docx";
            string inputDoc2 = MyDir + "Tables.docx";

            using (FileStream firstStreamIn = new FileStream(MyDir + "Big document.docx", FileMode.Open, FileAccess.Read))
            {
                using (FileStream secondStreamIn = new FileStream(MyDir + "Tables.docx", FileMode.Open, FileAccess.Read))
                {
                    OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
                    saveOptions.Password = "Aspose.Words";
                    using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MergeStreamContextDocuments.1.docx", FileMode.Create, FileAccess.ReadWrite))
                    {
                        MergerContext context = new MergerContext();
                        context.MergeFormatMode = MergeFormatMode.KeepSourceFormatting;
                        Merger.Create(context)
                        .From(firstStreamIn)
                        .From(secondStreamIn)
                        .To(streamOut, saveOptions)
                        .Execute();
                    }

                    LoadOptions firstLoadOptions = new LoadOptions();
                    firstLoadOptions.IgnoreOleData = true;
                    LoadOptions secondLoadOptions = new LoadOptions();
                    secondLoadOptions.IgnoreOleData = false;
                    using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MergeStreamContextDocuments.2.docx", FileMode.Create, FileAccess.ReadWrite))
                    {
                        MergerContext context2 = new MergerContext();
                        context2.MergeFormatMode = MergeFormatMode.KeepSourceFormatting;
                        Merger.Create(context2)
                        .From(firstStreamIn, firstLoadOptions)
                        .From(secondStreamIn, secondLoadOptions)
                        .To(streamOut, SaveFormat.Docx)
                        .Execute();
                    }
                }
            }
            //ExEnd:MergeStreamContextDocuments
        }

        [Test]
        public void MergeDocumentInstances()
        {
            //ExStart:MergeDocumentInstances
            //GistId:e386727403c2341ce4018bca370a5b41
            //ExFor:Merger.Merge(Document[], MergeFormatMode)
            //ExSummary:Shows how to merge input documents to a single document instance.
            DocumentBuilder firstDoc = new DocumentBuilder();
            firstDoc.Font.Size = 16;
            firstDoc.Font.Color = Color.Blue;
            firstDoc.Write("Hello first word!");

            DocumentBuilder secondDoc = new DocumentBuilder();
            secondDoc.Write("Hello second word!");

            Document mergedDoc = Merger.Merge(new Document[] { firstDoc.Document, secondDoc.Document }, MergeFormatMode.KeepSourceLayout);
            Assert.AreEqual("Hello first word!\fHello second word!\f", mergedDoc.GetText());
            //ExEnd:MergeDocumentInstances
        }

        [Test]
        public void Convert()
        {
            //ExStart:Convert
            //GistId:708ce40a68fac5003d46f6b4acfd5ff1
            //ExFor:Converter.Convert(String, String)
            //ExFor:Converter.Convert(String, String, SaveFormat)
            //ExFor:Converter.Convert(String, String, SaveOptions)
            //ExFor:Converter.Convert(String, LoadOptions, String, SaveOptions)
            //ExSummary:Shows how to convert documents with a single line of code.
            string doc = MyDir + "Document.docx";

            Converter.Convert(doc, ArtifactsDir + "LowCode.Convert.pdf");

            Converter.Convert(doc, ArtifactsDir + "LowCode.Convert.SaveFormat.rtf", SaveFormat.Rtf);

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.Password = "Aspose.Words";
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.IgnoreOleData = true;
            Converter.Convert(doc, loadOptions, ArtifactsDir + "LowCode.Convert.LoadOptions.docx", saveOptions);

            Converter.Convert(doc, ArtifactsDir + "LowCode.Convert.SaveOptions.docx", saveOptions);
            //ExEnd:Convert
        }

        [Test]
        public void ConvertContext()
        {
            //ExStart:ConvertContext
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Processor
            //ExFor:Processor.From(String, LoadOptions)
            //ExFor:Processor.To(String, SaveOptions)
            //ExFor:Processor.Execute
            //ExFor:Converter.Create(ConverterContext)
            //ExFor:ConverterContext
            //ExSummary:Shows how to convert documents with a single line of code using context.
            string doc = MyDir + "Big document.docx";

            Converter.Create(new ConverterContext())
                .From(doc)
                .To(ArtifactsDir + "LowCode.ConvertContext.1.pdf")
                .Execute();

            Converter.Create(new ConverterContext())
                .From(doc)
                .To(ArtifactsDir + "LowCode.ConvertContext.2.pdf", SaveFormat.Rtf)
                .Execute();

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.Password = "Aspose.Words";
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.IgnoreOleData = true;
            Converter.Create(new ConverterContext())
                .From(doc, loadOptions)
                .To(ArtifactsDir + "LowCode.ConvertContext.3.docx", saveOptions)
                .Execute();

            Converter.Create(new ConverterContext())
                .From(doc)
                .To(ArtifactsDir + "LowCode.ConvertContext.4.png", new ImageSaveOptions(SaveFormat.Png))
                .Execute();
            //ExEnd:ConvertContext
        }

        [Test]
        public void ConvertStream()
        {
            //ExStart:ConvertStream
            //GistId:708ce40a68fac5003d46f6b4acfd5ff1
            //ExFor:Converter.Convert(Stream, Stream, SaveFormat)
            //ExFor:Converter.Convert(Stream, Stream, SaveOptions)
            //ExFor:Converter.Convert(Stream, LoadOptions, Stream, SaveOptions)
            //ExSummary:Shows how to convert documents with a single line of code (Stream).
            using (FileStream streamIn = new FileStream(MyDir + "Big document.docx", FileMode.Open, FileAccess.Read))
            {
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.ConvertStream.1.docx", FileMode.Create, FileAccess.ReadWrite))
                    Converter.Convert(streamIn, streamOut, SaveFormat.Docx);

                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
                saveOptions.Password = "Aspose.Words";
                LoadOptions loadOptions = new LoadOptions();
                loadOptions.IgnoreOleData = true;
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.ConvertStream.2.docx", FileMode.Create, FileAccess.ReadWrite))
                    Converter.Convert(streamIn, loadOptions, streamOut, saveOptions);

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.ConvertStream.3.docx", FileMode.Create, FileAccess.ReadWrite))
                    Converter.Convert(streamIn, streamOut, saveOptions);
            }
            //ExEnd:ConvertStream
        }

        [Test]
        public void ConvertContextStream()
        {
            //ExStart:ConvertContextStream
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Processor
            //ExFor:Processor.From(Stream, LoadOptions)
            //ExFor:Processor.To(Stream, SaveFormat)
            //ExFor:Processor.To(Stream, SaveOptions)
            //ExFor:Processor.Execute
            //ExFor:Converter.Create(ConverterContext)
            //ExFor:ConverterContext
            //ExSummary:Shows how to convert documents from a stream with a single line of code using context.
            string doc = MyDir + "Document.docx";
            using (FileStream streamIn = new FileStream(MyDir + "Big document.docx", FileMode.Open, FileAccess.Read))
            {
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.ConvertContextStream.1.docx", FileMode.Create, FileAccess.ReadWrite))
                    Converter.Create(new ConverterContext())
                        .From(streamIn)
                        .To(streamOut, SaveFormat.Rtf)
                        .Execute();

                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
                saveOptions.Password = "Aspose.Words";
                LoadOptions loadOptions = new LoadOptions();
                loadOptions.IgnoreOleData = true;
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.ConvertContextStream.2.docx", FileMode.Create, FileAccess.ReadWrite))
                    Converter.Create(new ConverterContext())
                        .From(streamIn, loadOptions)
                        .To(streamOut, saveOptions)
                        .Execute();

                List<Stream> pages = new List<Stream>();
                Converter.Create(new ConverterContext())
                    .From(doc)
                    .To(pages, new ImageSaveOptions(SaveFormat.Png))
                    .Execute();
            }
            //ExEnd:ConvertContextStream
        }

        [Test]
        public void ConvertToImages()
        {
            //ExStart:ConvertToImages
            //GistId:708ce40a68fac5003d46f6b4acfd5ff1
            //ExFor:Converter.ConvertToImages(String, String)
            //ExFor:Converter.ConvertToImages(String, String, SaveFormat)
            //ExFor:Converter.ConvertToImages(String, String, ImageSaveOptions)
            //ExFor:Converter.ConvertToImages(String, LoadOptions, String, ImageSaveOptions)
            //ExSummary:Shows how to convert document to images.
            string doc = MyDir + "Big document.docx";

            Converter.Convert(doc, ArtifactsDir + "LowCode.ConvertToImages.1.png");

            Converter.Convert(doc, ArtifactsDir + "LowCode.ConvertToImages.2.jpeg", SaveFormat.Jpeg);

            LoadOptions loadOptions = new LoadOptions();
            loadOptions.IgnoreOleData = false;
            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Png);
            imageSaveOptions.PageSet = new PageSet(1);
            Converter.Convert(doc, loadOptions, ArtifactsDir + "LowCode.ConvertToImages.3.png", imageSaveOptions);

            Converter.Convert(doc, ArtifactsDir + "LowCode.ConvertToImages.4.png", imageSaveOptions);
            //ExEnd:ConvertToImages
        }

        [Test]
        public void ConvertToImagesStream()
        {
            //ExStart:ConvertToImagesStream
            //GistId:708ce40a68fac5003d46f6b4acfd5ff1
            //ExFor:Converter.ConvertToImages(String, SaveFormat)
            //ExFor:Converter.ConvertToImages(String, ImageSaveOptions)
            //ExFor:Converter.ConvertToImages(Document, SaveFormat)
            //ExFor:Converter.ConvertToImages(Document, ImageSaveOptions)
            //ExSummary:Shows how to convert document to images stream.
            string doc = MyDir + "Big document.docx";

            Stream[] streams = Converter.ConvertToImages(doc, SaveFormat.Png);

            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Png);
            imageSaveOptions.PageSet = new PageSet(1);
            streams = Converter.ConvertToImages(doc, imageSaveOptions);

            streams = Converter.ConvertToImages(new Document(doc), SaveFormat.Png);

            streams = Converter.ConvertToImages(new Document(doc), imageSaveOptions);
            //ExEnd:ConvertToImagesStream
        }

        [Test]
        public void ConvertToImagesFromStream()
        {
            //ExStart:ConvertToImagesFromStream
            //GistId:708ce40a68fac5003d46f6b4acfd5ff1
            //ExFor:Converter.ConvertToImages(Stream, SaveFormat)
            //ExFor:Converter.ConvertToImages(Stream, ImageSaveOptions)
            //ExFor:Converter.ConvertToImages(Stream, LoadOptions, ImageSaveOptions)
            //ExSummary:Shows how to convert document to images from stream.
            using (FileStream streamIn = new FileStream(MyDir + "Big document.docx", FileMode.Open, FileAccess.Read))
            {
                Stream[] streams = Converter.ConvertToImages(streamIn, SaveFormat.Jpeg);

                ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Png);
                imageSaveOptions.PageSet = new PageSet(1);
                streams = Converter.ConvertToImages(streamIn, imageSaveOptions);

                LoadOptions loadOptions = new LoadOptions();
                loadOptions.IgnoreOleData = false;
                Converter.ConvertToImages(streamIn, loadOptions, imageSaveOptions);
            }
            //ExEnd:ConvertToImagesFromStream
        }

        [TestCase("Protected pdf document.pdf", "PDF")]
        [TestCase("Pdf Document.pdf", "HTML")]
        [TestCase("Pdf Document.pdf", "XPS")]
        [TestCase("Images.pdf", "JPEG")]
        [TestCase("Images.pdf", "PNG")]
        [TestCase("Images.pdf", "TIFF")]
        [TestCase("Images.pdf", "BMP")]
        public void PdfRenderer(string docName, string format)
        {
            switch (format)
            {
                case "PDF":
                    LoadOptions loadOptions = new LoadOptions();
                    loadOptions.Password = "{Asp0se}P@ssw0rd";
                    SaveTo(docName, loadOptions, new PdfSaveOptions(), "pdf");
                    AssertResult("pdf");

                    break;

                case "HTML":
                    HtmlFixedSaveOptions htmlSaveOptions = new HtmlFixedSaveOptions();
                    htmlSaveOptions.PageSet = new PageSet(0);
                    htmlSaveOptions.PrettyFormat = true;
                    htmlSaveOptions.ExportEmbeddedFonts = true;
                    htmlSaveOptions.ExportEmbeddedCss = true;
                    SaveTo(docName, new LoadOptions(), htmlSaveOptions, "html");
                    AssertResult("html");

                    break;

                case "XPS":
                    SaveTo(docName, new LoadOptions(), new XpsSaveOptions(), "xps");
                    AssertResult("xps");

                    break;

                case "JPEG":
                    ImageSaveOptions jpegSaveOptions = new ImageSaveOptions(SaveFormat.Jpeg);
                    jpegSaveOptions.JpegQuality = 10;
                    SaveTo(docName, new LoadOptions(), jpegSaveOptions, "jpeg");
                    AssertResult("jpeg");

                    break;

                case "PNG":
                    ImageSaveOptions pngSaveOptions = new ImageSaveOptions(SaveFormat.Png);
                    pngSaveOptions.PageSet = new PageSet(0, 1);
                    pngSaveOptions.JpegQuality = 50;
                    SaveTo(docName, new LoadOptions(), pngSaveOptions, "png");
                    AssertResult("png");

                    break;

                case "TIFF":
                    ImageSaveOptions tiffSaveOptions = new ImageSaveOptions(SaveFormat.Tiff);
                    tiffSaveOptions.JpegQuality = 100;
                    SaveTo(docName, new LoadOptions(), tiffSaveOptions, "tiff");
                    AssertResult("tiff");

                    break;

                case "BMP":
                    ImageSaveOptions bmpSaveOptions = new ImageSaveOptions(SaveFormat.Bmp);
                    SaveTo(docName, new LoadOptions(), bmpSaveOptions, "bmp");
                    AssertResult("bmp");

                    break;
            }
        }

        private void SaveTo(string docName, LoadOptions loadOptions, SaveOptions saveOptions, string fileExt)
        {
            using (var pdfDoc = File.OpenRead(MyDir + docName))
            {
                MemoryStream stream = new MemoryStream();
                IReadOnlyList<Stream> imagesStream = new List<Stream>();

                if (fileExt == "pdf")
                {
                    Converter.Convert(pdfDoc, loadOptions, stream, saveOptions);
                }
                else if (fileExt == "html")
                {
                    Converter.Convert(pdfDoc, loadOptions, stream, saveOptions);
                }
                else if (fileExt == "xps")
                {
                    Converter.Convert(pdfDoc, loadOptions, stream, saveOptions);
                }
                else if (fileExt == "jpeg" || fileExt == "png" || fileExt == "tiff" || fileExt == "bmp")
                {
                    imagesStream = Converter.ConvertToImages(pdfDoc, loadOptions, (ImageSaveOptions)saveOptions);
                }

                stream.Position = 0;
                if (imagesStream.Count != 0)
                {
                    for (int i = 0; i < imagesStream.Count; i++)
                    {
                        using (FileStream resultDoc = new FileStream(ArtifactsDir + string.Format("PdfRenderer_{0}.{1}", i, fileExt), FileMode.Create))
                            imagesStream[i].CopyTo(resultDoc);
                    }
                }
                else
                {
                    using (FileStream resultDoc = new FileStream(ArtifactsDir + string.Format("PdfRenderer.{0}", fileExt), FileMode.Create))
                        stream.CopyTo(resultDoc);
                }
            }
        }

        private void AssertResult(string fileExt)
        {
            if (fileExt == "jpeg" || fileExt == "png" || fileExt == "tiff" || fileExt == "bmp")
            {
                Regex reg = new Regex("PdfRenderer_*");

                var images = Directory.GetFiles(ArtifactsDir, string.Format("*.{0}", fileExt))
                                     .Where(path => reg.IsMatch(path))
                                     .ToList();

                if (fileExt == "png")
                    Assert.AreEqual(2, images.Count);
                else if (fileExt == "tiff")
                    Assert.AreEqual(1, images.Count);
                else
                    Assert.AreEqual(5, images.Count);
            }
            else
            {
                if (fileExt == "xps")
                {
                    var doc = new XpsDocument(ArtifactsDir + string.Format("PdfRenderer.{0}", fileExt));
                    AssertXpsText(doc);
                }
                else if (fileExt == "pdf")
                {
                    Document doc = new Document(ArtifactsDir + string.Format("PdfRenderer.{0}", fileExt));
                    var content = doc.GetText();
                    Console.WriteLine(content);
                    Assert.IsTrue(content.Contains("Heading 1.1.1.2"));
                }
                else
                {
                    var content = File.ReadAllText(ArtifactsDir + string.Format("PdfRenderer.{0}", fileExt));
                    Console.WriteLine(content);
                    Assert.IsTrue(content.Contains("Heading 1.1.1.2"));
                }
            }
        }

        private static void AssertXpsText(XpsDocument doc)
        {
            AssertXpsText(doc.SelectActivePage(1));
        }

        private static void AssertXpsText(XpsElement element)
        {
            for (int i = 0; i < element.Count; i++)
                AssertXpsText(element[i]);
            if (element is XpsGlyphs)
                Assert.IsTrue(new[] { "Heading 1", "Head", "ing 1" }.Any(c => ((XpsGlyphs)element).UnicodeString.Contains(c)));
        }

        [Test]
        public void CompareDocuments()
        {
            //ExStart:CompareDocuments
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Comparer.Compare(String, String, String, String, DateTime, CompareOptions)
            //ExFor:Comparer.Compare(String, String, String, SaveFormat, String, DateTime, CompareOptions)
            //ExSummary:Shows how to simple compare documents.
            // There is a several ways to compare documents:
            string firstDoc = MyDir + "Table column bookmarks.docx";
            string secondDoc = MyDir + "Table column bookmarks.doc";

            Comparer.Compare(firstDoc, secondDoc, ArtifactsDir + "LowCode.CompareDocuments.1.docx", "Author", new DateTime());
            Comparer.Compare(firstDoc, secondDoc, ArtifactsDir + "LowCode.CompareDocuments.2.docx", SaveFormat.Docx, "Author", new DateTime());

            CompareOptions compareOptions = new CompareOptions();
            compareOptions.IgnoreCaseChanges = true;
            Comparer.Compare(firstDoc, secondDoc, ArtifactsDir + "LowCode.CompareDocuments.3.docx", "Author", new DateTime(), compareOptions);
            Comparer.Compare(firstDoc, secondDoc, ArtifactsDir + "LowCode.CompareDocuments.4.docx", SaveFormat.Docx, "Author", new DateTime(), compareOptions);
            //ExEnd:CompareDocuments
        }

        [Test]
        public void CompareContextDocuments()
        {
            //ExStart:CompareContextDocuments
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Comparer.Create(ComparerContext)
            //ExFor:ComparerContext
            //ExFor:ComparerContext.CompareOptions
            //ExSummary:Shows how to simple compare documents using context.
            // There is a several ways to compare documents:
            string firstDoc = MyDir + "Table column bookmarks.docx";
            string secondDoc = MyDir + "Table column bookmarks.doc";

            ComparerContext comparerContext = new ComparerContext();
            comparerContext.CompareOptions.IgnoreCaseChanges = true;
            comparerContext.Author = "Author";
            comparerContext.DateTime = new DateTime();

            Comparer.Create(comparerContext)
                .From(firstDoc)
                .From(secondDoc)
                .To(ArtifactsDir + "LowCode.CompareContextDocuments.docx")
                .Execute();
            //ExEnd:CompareContextDocuments
        }

        [Test]
        public void CompareStreamDocuments()
        {
            //ExStart:CompareStreamDocuments
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Comparer.Compare(Stream, Stream, Stream, SaveFormat, String, DateTime, CompareOptions)
            //ExSummary:Shows how to compare documents from the stream.
            // There is a several ways to compare documents from the stream:
            using (FileStream firstStreamIn = new FileStream(MyDir + "Table column bookmarks.docx", FileMode.Open, FileAccess.Read))
            {
                using (FileStream secondStreamIn = new FileStream(MyDir + "Table column bookmarks.doc", FileMode.Open, FileAccess.Read))
                {
                    using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.CompareStreamDocuments.1.docx", FileMode.Create, FileAccess.ReadWrite))
                        Comparer.Compare(firstStreamIn, secondStreamIn, streamOut, SaveFormat.Docx, "Author", new DateTime());

                    using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.CompareStreamDocuments.2.docx", FileMode.Create, FileAccess.ReadWrite))
                    {
                        CompareOptions compareOptions = new CompareOptions();
                        compareOptions.IgnoreCaseChanges = true;
                        Comparer.Compare(firstStreamIn, secondStreamIn, streamOut, SaveFormat.Docx, "Author", new DateTime(), compareOptions);
                    }
                }
            }
            //ExEnd:CompareStreamDocuments
        }

        [Test]
        public void CompareContextStreamDocuments()
        {
            //ExStart:CompareContextStreamDocuments
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Comparer.Create(ComparerContext)
            //ExFor:ComparerContext
            //ExFor:ComparerContext.CompareOptions
            //ExSummary:Shows how to compare documents from the stream using context.
            // There is a several ways to compare documents from the stream:
            using (FileStream firstStreamIn = new FileStream(MyDir + "Table column bookmarks.docx", FileMode.Open, FileAccess.Read))
            {
                using (FileStream secondStreamIn = new FileStream(MyDir + "Table column bookmarks.doc", FileMode.Open, FileAccess.Read))
                {
                    ComparerContext comparerContext = new ComparerContext();
                    comparerContext.CompareOptions.IgnoreCaseChanges = true;
                    comparerContext.Author = "Author";
                    comparerContext.DateTime = new DateTime();

                    using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.CompareContextStreamDocuments.docx", FileMode.Create, FileAccess.ReadWrite))
                        Comparer.Create(comparerContext)
                            .From(firstStreamIn)
                            .From(secondStreamIn)
                            .To(streamOut, SaveFormat.Docx)
                            .Execute();
                }
            }
            //ExEnd:CompareContextStreamDocuments
        }

        [Test]
        public void CompareDocumentsToimages()
        {
            //ExStart:CompareDocumentsToimages
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Comparer.CompareToImages(Stream, Stream, ImageSaveOptions, String, DateTime, CompareOptions)
            //ExSummary:Shows how to compare documents and save results as images.
            // There is a several ways to compare documents:
            string firstDoc = MyDir + "Table column bookmarks.docx";
            string secondDoc = MyDir + "Table column bookmarks.doc";

            Stream[] pages = Comparer.CompareToImages(firstDoc, secondDoc, new ImageSaveOptions(SaveFormat.Png), "Author", new DateTime());

            using (FileStream firstStreamIn = new FileStream(firstDoc, FileMode.Open, FileAccess.Read))
            {
                using (FileStream secondStreamIn = new FileStream(secondDoc, FileMode.Open, FileAccess.Read))
                {
                    CompareOptions compareOptions = new CompareOptions();
                    compareOptions.IgnoreCaseChanges = true;
                    pages = Comparer.CompareToImages(firstStreamIn, secondStreamIn, new ImageSaveOptions(SaveFormat.Png), "Author", new DateTime(), compareOptions);
                }
            }
            //ExEnd:CompareDocumentsToimages
        }

        [Test]
        public void MailMerge()
        {
            //ExStart:MailMerge
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:MailMergeOptions
            //ExFor:MailMergeOptions.TrimWhitespaces
            //ExFor:MailMerger.Execute(String, String, String[], Object[])
            //ExFor:MailMerger.Execute(String, String, SaveFormat, String[], Object[], MailMergeOptions)
            //ExSummary:Shows how to do mail merge operation for a single record.
            // There is a several ways to do mail merge operation:
            string doc = MyDir + "Mail merge.doc";

            string[] fieldNames = new string[] { "FirstName", "Location", "SpecialCharsInName()" };
            string[] fieldValues = new string[] { "James Bond", "London", "Classified" };

            MailMerger.Execute(doc, ArtifactsDir + "LowCode.MailMerge.1.docx", fieldNames, fieldValues);
            MailMerger.Execute(doc, ArtifactsDir + "LowCode.MailMerge.2.docx", SaveFormat.Docx, fieldNames, fieldValues);
            MailMergeOptions mailMergeOptions = new MailMergeOptions();
            mailMergeOptions.TrimWhitespaces = true;
            MailMerger.Execute(doc, ArtifactsDir + "LowCode.MailMerge.3.docx", SaveFormat.Docx, fieldNames, fieldValues, mailMergeOptions);
            //ExEnd:MailMerge
        }

        [Test]
        public void MailMergeContext()
        {
            //ExStart:MailMergeContext
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:MailMerger.Create(MailMergerContext)
            //ExFor:MailMergerContext
            //ExFor:MailMergerContext.SetSimpleDataSource(String[], Object[])
            //ExFor:MailMergerContext.MailMergeOptions
            //ExSummary:Shows how to do mail merge operation for a single record using context.
            // There is a several ways to do mail merge operation:
            string doc = MyDir + "Mail merge.doc";

            string[] fieldNames = new string[] { "FirstName", "Location", "SpecialCharsInName()" };
            string[] fieldValues = new string[] { "James Bond", "London", "Classified" };

            MailMergerContext mailMergerContext = new MailMergerContext();
            mailMergerContext.SetSimpleDataSource(fieldNames, fieldValues);
            mailMergerContext.MailMergeOptions.TrimWhitespaces = true;

            MailMerger.Create(mailMergerContext)
                .From(doc)
                .To(ArtifactsDir + "LowCode.MailMergeContext.docx")
                .Execute();
            //ExEnd:MailMergeContext
        }

        [Test]
        public void MailMergeToImages()
        {
            //ExStart:MailMergeToImages
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:MailMerger.ExecuteToImages(String, ImageSaveOptions, String[], Object[], MailMergeOptions)
            //ExSummary:Shows how to do mail merge operation for a single record and save result to images.
            // There is a several ways to do mail merge operation:
            string doc = MyDir + "Mail merge.doc";

            string[] fieldNames = new string[] { "FirstName", "Location", "SpecialCharsInName()" };
            string[] fieldValues = new string[] { "James Bond", "London", "Classified" };

            Stream[] images = MailMerger.ExecuteToImages(doc, new ImageSaveOptions(SaveFormat.Png), fieldNames, fieldValues);
            MailMergeOptions mailMergeOptions = new MailMergeOptions();
            mailMergeOptions.TrimWhitespaces = true;
            images = MailMerger.ExecuteToImages(doc, new ImageSaveOptions(SaveFormat.Png), fieldNames, fieldValues, mailMergeOptions);
            //ExEnd:MailMergeToImages
        }

        [Test]
        public void MailMergeStream()
        {
            //ExStart:MailMergeStream
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, String[], Object[], MailMergeOptions)
            //ExSummary:Shows how to do mail merge operation for a single record from the stream.
            // There is a several ways to do mail merge operation using documents from the stream:
            string[] fieldNames = new string[] { "FirstName", "Location", "SpecialCharsInName()" };
            string[] fieldValues = new string[] { "James Bond", "London", "Classified" };

            using (FileStream streamIn = new FileStream(MyDir + "Mail merge.doc", FileMode.Open, FileAccess.Read))
            {
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MailMergeStream.1.docx", FileMode.Create, FileAccess.ReadWrite))
                    MailMerger.Execute(streamIn, streamOut, SaveFormat.Docx, fieldNames, fieldValues);

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MailMergeStream.2.docx", FileMode.Create, FileAccess.ReadWrite))
                {
                    MailMergeOptions mailMergeOptions = new MailMergeOptions();
                    mailMergeOptions.TrimWhitespaces = true;
                    MailMerger.Execute(streamIn, streamOut, SaveFormat.Docx, fieldNames, fieldValues, mailMergeOptions);
                }
            }
            //ExEnd:MailMergeStream
        }

        [Test]
        public void MailMergeContextStream()
        {
            //ExStart:MailMergeContextStream
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:MailMerger.Create(MailMergerContext)
            //ExFor:MailMergerContext
            //ExFor:MailMergerContext.SetSimpleDataSource(String[], Object[])
            //ExFor:MailMergerContext.MailMergeOptions
            //ExSummary:Shows how to do mail merge operation for a single record from the stream using context.
            // There is a several ways to do mail merge operation using documents from the stream:
            string[] fieldNames = new string[] { "FirstName", "Location", "SpecialCharsInName()" };
            string[] fieldValues = new string[] { "James Bond", "London", "Classified" };

            using (FileStream streamIn = new FileStream(MyDir + "Mail merge.doc", FileMode.Open, FileAccess.Read))
            {
                MailMergerContext mailMergerContext = new MailMergerContext();
                mailMergerContext.SetSimpleDataSource(fieldNames, fieldValues);
                mailMergerContext.MailMergeOptions.TrimWhitespaces = true;

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MailMergeContextStream.docx", FileMode.Create, FileAccess.ReadWrite))
                    MailMerger.Create(mailMergerContext)
                        .From(streamIn)
                        .To(streamOut, SaveFormat.Docx)
                        .Execute();
            }
            //ExEnd:MailMergeContextStream
        }

        [Test]
        public void MailMergeStreamToImages()
        {
            //ExStart:MailMergeStreamToImages
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:MailMerger.ExecuteToImages(Stream, ImageSaveOptions, String[], Object[], MailMergeOptions)
            //ExSummary:Shows how to do mail merge operation for a single record from the stream and save result to images.
            // There is a several ways to do mail merge operation using documents from the stream:
            string[] fieldNames = new string[] { "FirstName", "Location", "SpecialCharsInName()" };
            string[] fieldValues = new string[] { "James Bond", "London", "Classified" };

            using (FileStream streamIn = new FileStream(MyDir + "Mail merge.doc", FileMode.Open, FileAccess.Read))
            {
                Stream[] images = MailMerger.ExecuteToImages(streamIn, new ImageSaveOptions(SaveFormat.Png), fieldNames, fieldValues);

                MailMergeOptions mailMergeOptions = new MailMergeOptions();
                mailMergeOptions.TrimWhitespaces = true;
                images = MailMerger.ExecuteToImages(streamIn, new ImageSaveOptions(SaveFormat.Png), fieldNames, fieldValues, mailMergeOptions);
            }
            //ExEnd:MailMergeStreamToImages
        }

        [Test]
        public void MailMergeDataRow()
        {
            //ExStart:MailMergeDataRow
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:MailMerger.Execute(String, String, DataRow)
            //ExFor:MailMerger.Execute(String, String, SaveFormat, DataRow, MailMergeOptions)
            //ExSummary:Shows how to do mail merge operation from a DataRow.
            // There is a several ways to do mail merge operation from a DataRow:
            string doc = MyDir + "Mail merge.doc";

            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("FirstName");
            dataTable.Columns.Add("Location");
            dataTable.Columns.Add("SpecialCharsInName()");

            DataRow dataRow = dataTable.Rows.Add(new string[] { "James Bond", "London", "Classified" });

            MailMerger.Execute(doc, ArtifactsDir + "LowCode.MailMergeDataRow.1.docx", dataRow);
            MailMerger.Execute(doc, ArtifactsDir + "LowCode.MailMergeDataRow.2.docx", SaveFormat.Docx, dataRow);
            MailMergeOptions mailMergeOptions = new MailMergeOptions();
            mailMergeOptions.TrimWhitespaces = true;
            MailMerger.Execute(doc, ArtifactsDir + "LowCode.MailMergeDataRow.3.docx", SaveFormat.Docx, dataRow, mailMergeOptions);
            //ExEnd:MailMergeDataRow
        }

        [Test]
        public void MailMergeContextDataRow()
        {
            //ExStart:MailMergeContextDataRow
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:MailMerger.Create(MailMergerContext)
            //ExFor:MailMergerContext
            //ExFor:MailMergerContext.SetSimpleDataSource(DataRow)
            //ExSummary:Shows how to do mail merge operation from a DataRow using context.
            // There is a several ways to do mail merge operation from a DataRow:
            string doc = MyDir + "Mail merge.doc";

            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("FirstName");
            dataTable.Columns.Add("Location");
            dataTable.Columns.Add("SpecialCharsInName()");

            DataRow dataRow = dataTable.Rows.Add(new string[] { "James Bond", "London", "Classified" });

            MailMergerContext mailMergerContext = new MailMergerContext();
            mailMergerContext.SetSimpleDataSource(dataRow);
            mailMergerContext.MailMergeOptions.TrimWhitespaces = true;

            MailMerger.Create(mailMergerContext)
                .From(doc)
                .To(ArtifactsDir + "LowCode.MailMergeContextDataRow.docx")
                .Execute();
            //ExEnd:MailMergeContextDataRow
        }

        [Test]
        public void MailMergeToImagesDataRow()
        {
            //ExStart:MailMergeToImagesDataRow
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:MailMerger.ExecuteToImages(String, ImageSaveOptions, DataRow, MailMergeOptions)
            //ExSummary:Shows how to do mail merge operation from a DataRow and save result to images.
            // There is a several ways to do mail merge operation from a DataRow:
            string doc = MyDir + "Mail merge.doc";

            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("FirstName");
            dataTable.Columns.Add("Location");
            dataTable.Columns.Add("SpecialCharsInName()");

            DataRow dataRow = dataTable.Rows.Add(new string[] { "James Bond", "London", "Classified" });

            Stream[] images = MailMerger.ExecuteToImages(doc, new ImageSaveOptions(SaveFormat.Png), dataRow);
            MailMergeOptions mailMergeOptions = new MailMergeOptions();
            mailMergeOptions.TrimWhitespaces = true;
            images = MailMerger.ExecuteToImages(doc, new ImageSaveOptions(SaveFormat.Png), dataRow, mailMergeOptions);
            //ExEnd:MailMergeToImagesDataRow
        }

        [Test]
        public void MailMergeStreamDataRow()
        {
            //ExStart:MailMergeStreamDataRow
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, DataRow, MailMergeOptions)
            //ExSummary:Shows how to do mail merge operation from a DataRow using documents from the stream.
            // There is a several ways to do mail merge operation from a DataRow using documents from the stream:
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("FirstName");
            dataTable.Columns.Add("Location");
            dataTable.Columns.Add("SpecialCharsInName()");

            DataRow dataRow = dataTable.Rows.Add(new string[] { "James Bond", "London", "Classified" });

            using (FileStream streamIn = new FileStream(MyDir + "Mail merge.doc", FileMode.Open, FileAccess.Read))
            {
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MailMergeStreamDataRow.1.docx", FileMode.Create, FileAccess.ReadWrite))
                    MailMerger.Execute(streamIn, streamOut, SaveFormat.Docx, dataRow);

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MailMergeStreamDataRow.2.docx", FileMode.Create, FileAccess.ReadWrite))
                {
                    MailMergeOptions mailMergeOptions = new MailMergeOptions();
                    mailMergeOptions.TrimWhitespaces = true;
                    MailMerger.Execute(streamIn, streamOut, SaveFormat.Docx, dataRow, mailMergeOptions);
                }
            }
            //ExEnd:MailMergeStreamDataRow
        }

        [Test]
        public void MailMergeContextStreamDataRow()
        {
            //ExStart:MailMergeContextStreamDataRow
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:MailMerger.Create(MailMergerContext)
            //ExFor:MailMergerContext
            //ExFor:MailMergerContext.SetSimpleDataSource(DataRow)
            //ExSummary:Shows how to do mail merge operation from a DataRow using documents from the stream using context.
            // There is a several ways to do mail merge operation from a DataRow using documents from the stream:
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("FirstName");
            dataTable.Columns.Add("Location");
            dataTable.Columns.Add("SpecialCharsInName()");

            DataRow dataRow = dataTable.Rows.Add(new string[] { "James Bond", "London", "Classified" });

            using (FileStream streamIn = new FileStream(MyDir + "Mail merge.doc", FileMode.Open, FileAccess.Read))
            {
                MailMergerContext mailMergerContext = new MailMergerContext();
                mailMergerContext.SetSimpleDataSource(dataRow);
                mailMergerContext.MailMergeOptions.TrimWhitespaces = true;

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MailMergeContextStreamDataRow.docx", FileMode.Create, FileAccess.ReadWrite))
                    MailMerger.Create(mailMergerContext)
                        .From(streamIn)
                        .To(streamOut, SaveFormat.Docx)
                        .Execute();
            }
            //ExEnd:MailMergeContextStreamDataRow
        }

        [Test]
        public void MailMergeStreamToImagesDataRow()
        {
            //ExStart:MailMergeStreamToImagesDataRow
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:MailMerger.ExecuteToImages(Stream, ImageSaveOptions, DataRow, MailMergeOptions)
            //ExSummary:Shows how to do mail merge operation from a DataRow using documents from the stream and save result to images.
            // There is a several ways to do mail merge operation from a DataRow using documents from the stream:
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("FirstName");
            dataTable.Columns.Add("Location");
            dataTable.Columns.Add("SpecialCharsInName()");

            DataRow dataRow = dataTable.Rows.Add(new string[] { "James Bond", "London", "Classified" });

            using (FileStream streamIn = new FileStream(MyDir + "Mail merge.doc", FileMode.Open, FileAccess.Read))
            {
                Stream[] images = MailMerger.ExecuteToImages(streamIn, new ImageSaveOptions(SaveFormat.Png), dataRow);
                MailMergeOptions mailMergeOptions = new MailMergeOptions();
                mailMergeOptions.TrimWhitespaces = true;
                images = MailMerger.ExecuteToImages(streamIn, new ImageSaveOptions(SaveFormat.Png), dataRow, mailMergeOptions);
            }
            //ExEnd:MailMergeStreamToImagesDataRow
        }

        [Test]
        public void MailMergeDataTable()
        {
            //ExStart:MailMergeDataTable
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:MailMerger.Execute(String, String, DataTable)
            //ExFor:MailMerger.Execute(String, String, SaveFormat, DataTable, MailMergeOptions)
            //ExSummary:Shows how to do mail merge operation from a DataTable.
            // There is a several ways to do mail merge operation from a DataTable:
            string doc = MyDir + "Mail merge.doc";

            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("FirstName");
            dataTable.Columns.Add("Location");
            dataTable.Columns.Add("SpecialCharsInName()");

            DataRow dataRow = dataTable.Rows.Add(new string[] { "James Bond", "London", "Classified" });

            MailMerger.Execute(doc, ArtifactsDir + "LowCode.MailMergeDataTable.1.docx", dataTable);
            MailMerger.Execute(doc, ArtifactsDir + "LowCode.MailMergeDataTable.2.docx", SaveFormat.Docx, dataTable);
            MailMergeOptions mailMergeOptions = new MailMergeOptions();
            mailMergeOptions.TrimWhitespaces = true;
            MailMerger.Execute(doc, ArtifactsDir + "LowCode.MailMergeDataTable.3.docx", SaveFormat.Docx, dataTable, mailMergeOptions);
            //ExEnd:MailMergeDataTable
        }

        [Test]
        public void MailMergeContextDataTable()
        {
            //ExStart:MailMergeContextDataTable
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:MailMerger.Create(MailMergerContext)
            //ExFor:MailMergerContext
            //ExFor:MailMergerContext.SetSimpleDataSource(DataTable)
            //ExSummary:Shows how to do mail merge operation from a DataTable using context.
            // There is a several ways to do mail merge operation from a DataTable:
            string doc = MyDir + "Mail merge.doc";

            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("FirstName");
            dataTable.Columns.Add("Location");
            dataTable.Columns.Add("SpecialCharsInName()");

            DataRow dataRow = dataTable.Rows.Add(new string[] { "James Bond", "London", "Classified" });

            MailMergerContext mailMergerContext = new MailMergerContext();
            mailMergerContext.SetSimpleDataSource(dataTable);
            mailMergerContext.MailMergeOptions.TrimWhitespaces = true;

            MailMerger.Create(mailMergerContext)
                .From(doc)
                .To(ArtifactsDir + "LowCode.MailMergeContextDataTable.docx")
                .Execute();
            //ExEnd:MailMergeContextDataTable
        }

        [Test]
        public void MailMergeToImagesDataTable()
        {
            //ExStart:MailMergeToImagesDataTable
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:MailMerger.ExecuteToImages(String, ImageSaveOptions, DataTable, MailMergeOptions)
            //ExSummary:Shows how to do mail merge operation from a DataTable and save result to images.
            // There is a several ways to do mail merge operation from a DataTable:
            string doc = MyDir + "Mail merge.doc";

            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("FirstName");
            dataTable.Columns.Add("Location");
            dataTable.Columns.Add("SpecialCharsInName()");

            DataRow dataRow = dataTable.Rows.Add(new string[] { "James Bond", "London", "Classified" });

            Stream[] images = MailMerger.ExecuteToImages(doc, new ImageSaveOptions(SaveFormat.Png), dataTable);
            MailMergeOptions mailMergeOptions = new MailMergeOptions();
            mailMergeOptions.TrimWhitespaces = true;
            images = MailMerger.ExecuteToImages(doc, new ImageSaveOptions(SaveFormat.Png), dataTable, mailMergeOptions);
            //ExEnd:MailMergeToImagesDataTable
        }

        [Test]
        public void MailMergeStreamDataTable()
        {
            //ExStart:MailMergeStreamDataTable
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, DataTable, MailMergeOptions)
            //ExSummary:Shows how to do mail merge operation from a DataTable using documents from the stream.
            // There is a several ways to do mail merge operation from a DataTable using documents from the stream:
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("FirstName");
            dataTable.Columns.Add("Location");
            dataTable.Columns.Add("SpecialCharsInName()");

            DataRow dataRow = dataTable.Rows.Add(new string[] { "James Bond", "London", "Classified" });

            using (FileStream streamIn = new FileStream(MyDir + "Mail merge.doc", FileMode.Open, FileAccess.Read))
            {
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MailMergeDataTable.1.docx", FileMode.Create, FileAccess.ReadWrite))
                    MailMerger.Execute(streamIn, streamOut, SaveFormat.Docx, dataTable);

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MailMergeDataTable.2.docx", FileMode.Create, FileAccess.ReadWrite))
                {
                    MailMergeOptions mailMergeOptions = new MailMergeOptions();
                    mailMergeOptions.TrimWhitespaces = true;
                    MailMerger.Execute(streamIn, streamOut, SaveFormat.Docx, dataTable, mailMergeOptions);
                }
            }
            //ExEnd:MailMergeStreamDataTable
        }

        [Test]
        public void MailMergeContextStreamDataTable()
        {
            //ExStart:MailMergeContextStreamDataTable
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Processor
            //ExFor:MailMerger.Create(MailMergerContext)
            //ExFor:MailMergerContext
            //ExFor:MailMergerContext.SetSimpleDataSource(DataTable)
            //ExSummary:Shows how to do mail merge operation from a DataTable using documents from the stream using context.
            // There is a several ways to do mail merge operation from a DataTable using documents from the stream:
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("FirstName");
            dataTable.Columns.Add("Location");
            dataTable.Columns.Add("SpecialCharsInName()");

            DataRow dataRow = dataTable.Rows.Add(new string[] { "James Bond", "London", "Classified" });

            using (FileStream streamIn = new FileStream(MyDir + "Mail merge.doc", FileMode.Open, FileAccess.Read))
            {
                MailMergerContext mailMergerContext = new MailMergerContext();
                mailMergerContext.SetSimpleDataSource(dataTable);
                mailMergerContext.MailMergeOptions.TrimWhitespaces = true;

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MailMergeContextStreamDataTable.docx", FileMode.Create, FileAccess.ReadWrite))
                    MailMerger.Create(mailMergerContext)
                        .From(streamIn)
                        .To(streamOut, SaveFormat.Docx)
                        .Execute();
            }
            //ExEnd:MailMergeContextStreamDataTable
        }

        [Test]
        public void MailMergeStreamToImagesDataTable()
        {
            //ExStart:MailMergeStreamToImagesDataTable
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:MailMerger.ExecuteToImages(Stream, ImageSaveOptions, DataTable, MailMergeOptions)
            //ExSummary:Shows how to do mail merge operation from a DataTable using documents from the stream and save to images.
            // There is a several ways to do mail merge operation from a DataTable using documents from the stream and save result to images:
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("FirstName");
            dataTable.Columns.Add("Location");
            dataTable.Columns.Add("SpecialCharsInName()");

            DataRow dataRow = dataTable.Rows.Add(new string[] { "James Bond", "London", "Classified" });

            using (FileStream streamIn = new FileStream(MyDir + "Mail merge.doc", FileMode.Open, FileAccess.Read))
            {
                Stream[] images = MailMerger.ExecuteToImages(streamIn, new ImageSaveOptions(SaveFormat.Png), dataTable);
                MailMergeOptions mailMergeOptions = new MailMergeOptions();
                mailMergeOptions.TrimWhitespaces = true;
                images = MailMerger.ExecuteToImages(streamIn, new ImageSaveOptions(SaveFormat.Png), dataTable, mailMergeOptions);
            }
            //ExEnd:MailMergeStreamToImagesDataTable
        }

        [Test]
        public void MailMergeWithRegionsDataTable()
        {
            //ExStart:MailMergeWithRegionsDataTable
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:MailMerger.ExecuteWithRegions(String, String, DataTable)
            //ExFor:MailMerger.ExecuteWithRegions(String, String, SaveFormat, DataTable, MailMergeOptions)
            //ExSummary:Shows how to do mail merge with regions operation from a DataTable.
            // There is a several ways to do mail merge with regions operation from a DataTable:
            string doc = MyDir + "Mail merge with regions.docx";

            DataTable dataTable = new DataTable("MyTable");
            dataTable.Columns.Add("FirstName");
            dataTable.Columns.Add("LastName");
            dataTable.Rows.Add(new object[] { "John", "Doe" });
            dataTable.Rows.Add(new object[] { "", "" });
            dataTable.Rows.Add(new object[] { "Jane", "Doe" });

            MailMerger.ExecuteWithRegions(doc, ArtifactsDir + "LowCode.MailMergeWithRegionsDataTable.1.docx", dataTable);
            MailMerger.ExecuteWithRegions(doc, ArtifactsDir + "LowCode.MailMergeWithRegionsDataTable.2.docx", SaveFormat.Docx, dataTable);
            MailMergeOptions mailMergeOptions = new MailMergeOptions();
            mailMergeOptions.TrimWhitespaces = true;
            MailMerger.ExecuteWithRegions(doc, ArtifactsDir + "LowCode.MailMergeWithRegionsDataTable.3.docx", SaveFormat.Docx, dataTable, mailMergeOptions);
            //ExEnd:MailMergeWithRegionsDataTable
        }

        [Test]
        public void MailMergeContextWithRegionsDataTable()
        {
            //ExStart:MailMergeContextWithRegionsDataTable
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:MailMerger.Create(MailMergerContext)
            //ExFor:MailMergerContext
            //ExFor:MailMergerContext.SetRegionsDataSource(DataTable)
            //ExSummary:Shows how to do mail merge with regions operation from a DataTable using context.
            // There is a several ways to do mail merge with regions operation from a DataTable:
            string doc = MyDir + "Mail merge with regions.docx";

            DataTable dataTable = new DataTable("MyTable");
            dataTable.Columns.Add("FirstName");
            dataTable.Columns.Add("LastName");
            dataTable.Rows.Add(new object[] { "John", "Doe" });
            dataTable.Rows.Add(new object[] { "", "" });
            dataTable.Rows.Add(new object[] { "Jane", "Doe" });

            MailMergerContext mailMergerContext = new MailMergerContext();
            mailMergerContext.SetRegionsDataSource(dataTable);
            mailMergerContext.MailMergeOptions.TrimWhitespaces = true;

            MailMerger.Create(mailMergerContext)
                .From(doc)
                .To(ArtifactsDir + "LowCode.MailMergeContextWithRegionsDataTable.docx")
                .Execute();
            //ExEnd:MailMergeContextWithRegionsDataTable
        }

        [Test]
        public void MailMergeWithRegionsToImagesDataTable()
        {
            //ExStart:MailMergeWithRegionsToImagesDataTable
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:MailMerger.ExecuteWithRegionsToImages(String, ImageSaveOptions, DataTable, MailMergeOptions)
            //ExSummary:Shows how to do mail merge with regions operation from a DataTable and save result to images.
            // There is a several ways to do mail merge with regions operation from a DataTable:
            string doc = MyDir + "Mail merge with regions.docx";

            DataTable dataTable = new DataTable("MyTable");
            dataTable.Columns.Add("FirstName");
            dataTable.Columns.Add("LastName");
            dataTable.Rows.Add(new object[] { "John", "Doe" });
            dataTable.Rows.Add(new object[] { "", "" });
            dataTable.Rows.Add(new object[] { "Jane", "Doe" });

            Stream[] images = MailMerger.ExecuteWithRegionsToImages(doc, new ImageSaveOptions(SaveFormat.Png), dataTable);
            MailMergeOptions mailMergeOptions = new MailMergeOptions();
            mailMergeOptions.TrimWhitespaces = true;
            images = MailMerger.ExecuteWithRegionsToImages(doc, new ImageSaveOptions(SaveFormat.Png), dataTable, mailMergeOptions);
            //ExEnd:MailMergeWithRegionsToImagesDataTable
        }

        [Test]
        public void MailMergeStreamWithRegionsDataTable()
        {
            //ExStart:MailMergeStreamWithRegionsDataTable
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:MailMerger.ExecuteWithRegions(Stream, Stream, SaveFormat, DataTable, MailMergeOptions)
            //ExSummary:Shows how to do mail merge with regions operation from a DataTable using documents from the stream.
            // There is a several ways to do mail merge with regions operation from a DataTable using documents from the stream:
            DataTable dataTable = new DataTable("MyTable");
            dataTable.Columns.Add("FirstName");
            dataTable.Columns.Add("LastName");
            dataTable.Rows.Add(new object[] { "John", "Doe" });
            dataTable.Rows.Add(new object[] { "", "" });
            dataTable.Rows.Add(new object[] { "Jane", "Doe" });

            using (FileStream streamIn = new FileStream(MyDir + "Mail merge.doc", FileMode.Open, FileAccess.Read))
            {
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MailMergeStreamWithRegionsDataTable.1.docx", FileMode.Create, FileAccess.ReadWrite))
                    MailMerger.ExecuteWithRegions(streamIn, streamOut, SaveFormat.Docx, dataTable);

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MailMergeStreamWithRegionsDataTable.2.docx", FileMode.Create, FileAccess.ReadWrite))
                {
                    MailMergeOptions mailMergeOptions = new MailMergeOptions();
                    mailMergeOptions.TrimWhitespaces = true;
                    MailMerger.ExecuteWithRegions(streamIn, streamOut, SaveFormat.Docx, dataTable, mailMergeOptions);
                }
            }
            //ExEnd:MailMergeStreamWithRegionsDataTable
        }

        [Test]
        public void MailMergeContextStreamWithRegionsDataTable()
        {
            //ExStart:MailMergeContextStreamWithRegionsDataTable
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:MailMerger.Create(MailMergerContext)
            //ExFor:MailMergerContext
            //ExFor:MailMergerContext.SetRegionsDataSource(DataTable)
            //ExSummary:Shows how to do mail merge with regions operation from a DataTable using documents from the stream using context.
            // There is a several ways to do mail merge with regions operation from a DataTable using documents from the stream:
            DataTable dataTable = new DataTable("MyTable");
            dataTable.Columns.Add("FirstName");
            dataTable.Columns.Add("LastName");
            dataTable.Rows.Add(new object[] { "John", "Doe" });
            dataTable.Rows.Add(new object[] { "", "" });
            dataTable.Rows.Add(new object[] { "Jane", "Doe" });

            using (FileStream streamIn = new FileStream(MyDir + "Mail merge.doc", FileMode.Open, FileAccess.Read))
            {
                MailMergerContext mailMergerContext = new MailMergerContext();
                mailMergerContext.SetRegionsDataSource(dataTable);
                mailMergerContext.MailMergeOptions.TrimWhitespaces = true;

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MailMergeContextStreamWithRegionsDataTable.docx", FileMode.Create, FileAccess.ReadWrite))
                    MailMerger.Create(mailMergerContext)
                        .From(streamIn)
                        .To(streamOut, SaveFormat.Docx)
                        .Execute();
            }
            //ExEnd:MailMergeContextStreamWithRegionsDataTable
        }

        [Test]
        public void MailMergeStreamWithRegionsToImagesDataTable()
        {
            //ExStart:MailMergeStreamWithRegionsToImagesDataTable
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:MailMerger.ExecuteWithRegionsToImages(Stream, ImageSaveOptions, DataTable, MailMergeOptions)
            //ExSummary:Shows how to do mail merge with regions operation from a DataTable using documents from the stream and save result to images.
            // There is a several ways to do mail merge with regions operation from a DataTable using documents from the stream:
            DataTable dataTable = new DataTable("MyTable");
            dataTable.Columns.Add("FirstName");
            dataTable.Columns.Add("LastName");
            dataTable.Rows.Add(new object[] { "John", "Doe" });
            dataTable.Rows.Add(new object[] { "", "" });
            dataTable.Rows.Add(new object[] { "Jane", "Doe" });

            using (FileStream streamIn = new FileStream(MyDir + "Mail merge.doc", FileMode.Open, FileAccess.Read))
            {
                Stream[] images = MailMerger.ExecuteWithRegionsToImages(streamIn, new ImageSaveOptions(SaveFormat.Png), dataTable);
                MailMergeOptions mailMergeOptions = new MailMergeOptions();
                mailMergeOptions.TrimWhitespaces = true;
                images = MailMerger.ExecuteWithRegionsToImages(streamIn, new ImageSaveOptions(SaveFormat.Png), dataTable, mailMergeOptions);
            }
            //ExEnd:MailMergeStreamWithRegionsToImagesDataTable
        }

        [Test]
        public void MailMergeWithRegionsDataSet()
        {
            //ExStart:MailMergeWithRegionsDataSet
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:MailMerger.ExecuteWithRegions(String, String, DataSet)
            //ExFor:MailMerger.ExecuteWithRegions(String, String, SaveFormat, DataSet, MailMergeOptions)
            //ExSummary:Shows how to do mail merge with regions operation from a DataSet.
            // There is a several ways to do mail merge with regions operation from a DataSet:
            string doc = MyDir + "Mail merge with regions data set.docx";

            DataTable tableCustomers = new DataTable("Customers");
            tableCustomers.Columns.Add("CustomerID");
            tableCustomers.Columns.Add("CustomerName");
            tableCustomers.Rows.Add(new object[] { 1, "John Doe" });
            tableCustomers.Rows.Add(new object[] { 2, "Jane Doe" });

            DataTable tableOrders = new DataTable("Orders");
            tableOrders.Columns.Add("CustomerID");
            tableOrders.Columns.Add("ItemName");
            tableOrders.Columns.Add("Quantity");
            tableOrders.Rows.Add(new object[] { 1, "Hawaiian", 2 });
            tableOrders.Rows.Add(new object[] { 2, "Pepperoni", 1 });
            tableOrders.Rows.Add(new object[] { 2, "Chicago", 1 });

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(tableCustomers);
            dataSet.Tables.Add(tableOrders);
            dataSet.Relations.Add(tableCustomers.Columns["CustomerID"], tableOrders.Columns["CustomerID"]);

            MailMerger.ExecuteWithRegions(doc, ArtifactsDir + "LowCode.MailMergeWithRegionsDataSet.1.docx", dataSet);
            MailMerger.ExecuteWithRegions(doc, ArtifactsDir + "LowCode.MailMergeWithRegionsDataSet.2.docx", SaveFormat.Docx, dataSet);
            MailMergeOptions mailMergeOptions = new MailMergeOptions();
            mailMergeOptions.TrimWhitespaces = true;
            MailMerger.ExecuteWithRegions(doc, ArtifactsDir + "LowCode.MailMergeWithRegionsDataSet.3.docx", SaveFormat.Docx, dataSet, mailMergeOptions);
            //ExEnd:MailMergeWithRegionsDataSet
        }

        [Test]
        public void MailMergeContextWithRegionsDataSet()
        {
            //ExStart:MailMergeContextWithRegionsDataSet
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:MailMerger.Create(MailMergerContext)
            //ExFor:MailMergerContext
            //ExFor:MailMergerContext.SetRegionsDataSource(DataSet)
            //ExSummary:Shows how to do mail merge with regions operation from a DataSet using context.
            // There is a several ways to do mail merge with regions operation from a DataSet:
            string doc = MyDir + "Mail merge with regions data set.docx";

            DataTable tableCustomers = new DataTable("Customers");
            tableCustomers.Columns.Add("CustomerID");
            tableCustomers.Columns.Add("CustomerName");
            tableCustomers.Rows.Add(new object[] { 1, "John Doe" });
            tableCustomers.Rows.Add(new object[] { 2, "Jane Doe" });

            DataTable tableOrders = new DataTable("Orders");
            tableOrders.Columns.Add("CustomerID");
            tableOrders.Columns.Add("ItemName");
            tableOrders.Columns.Add("Quantity");
            tableOrders.Rows.Add(new object[] { 1, "Hawaiian", 2 });
            tableOrders.Rows.Add(new object[] { 2, "Pepperoni", 1 });
            tableOrders.Rows.Add(new object[] { 2, "Chicago", 1 });

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(tableCustomers);
            dataSet.Tables.Add(tableOrders);
            dataSet.Relations.Add(tableCustomers.Columns["CustomerID"], tableOrders.Columns["CustomerID"]);

            MailMergerContext mailMergerContext = new MailMergerContext();
            mailMergerContext.SetRegionsDataSource(dataSet);
            mailMergerContext.MailMergeOptions.TrimWhitespaces = true;

            MailMerger.Create(mailMergerContext)
                .From(doc)
                .To(ArtifactsDir + "LowCode.MailMergeContextWithRegionsDataTable.docx")
                .Execute();
            //ExEnd:MailMergeContextWithRegionsDataSet
        }

        [Test]
        public void MailMergeWithRegionsToImagesDataSet()
        {
            //ExStart:MailMergeWithRegionsToImagesDataSet
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:MailMerger.ExecuteWithRegionsToImages(String, ImageSaveOptions, DataSet, MailMergeOptions)
            //ExSummary:Shows how to do mail merge with regions operation from a DataSet and save result to images.
            // There is a several ways to do mail merge with regions operation from a DataSet:
            string doc = MyDir + "Mail merge with regions data set.docx";

            DataTable tableCustomers = new DataTable("Customers");
            tableCustomers.Columns.Add("CustomerID");
            tableCustomers.Columns.Add("CustomerName");
            tableCustomers.Rows.Add(new object[] { 1, "John Doe" });
            tableCustomers.Rows.Add(new object[] { 2, "Jane Doe" });

            DataTable tableOrders = new DataTable("Orders");
            tableOrders.Columns.Add("CustomerID");
            tableOrders.Columns.Add("ItemName");
            tableOrders.Columns.Add("Quantity");
            tableOrders.Rows.Add(new object[] { 1, "Hawaiian", 2 });
            tableOrders.Rows.Add(new object[] { 2, "Pepperoni", 1 });
            tableOrders.Rows.Add(new object[] { 2, "Chicago", 1 });

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(tableCustomers);
            dataSet.Tables.Add(tableOrders);
            dataSet.Relations.Add(tableCustomers.Columns["CustomerID"], tableOrders.Columns["CustomerID"]);

            Stream[] images = MailMerger.ExecuteWithRegionsToImages(doc, new ImageSaveOptions(SaveFormat.Png), dataSet);
            MailMergeOptions mailMergeOptions = new MailMergeOptions();
            mailMergeOptions.TrimWhitespaces = true;
            images = MailMerger.ExecuteWithRegionsToImages(doc, new ImageSaveOptions(SaveFormat.Png), dataSet, mailMergeOptions);
            //ExEnd:MailMergeWithRegionsToImagesDataSet
        }

        [Test]
        public void MailMergeStreamWithRegionsDataSet()
        {
            //ExStart:MailMergeStreamWithRegionsDataSet
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:MailMerger.ExecuteWithRegions(Stream, Stream, SaveFormat, DataSet, MailMergeOptions)
            //ExSummary:Shows how to do mail merge with regions operation from a DataSet using documents from the stream.
            // There is a several ways to do mail merge with regions operation from a DataSet using documents from the stream:
            DataTable tableCustomers = new DataTable("Customers");
            tableCustomers.Columns.Add("CustomerID");
            tableCustomers.Columns.Add("CustomerName");
            tableCustomers.Rows.Add(new object[] { 1, "John Doe" });
            tableCustomers.Rows.Add(new object[] { 2, "Jane Doe" });

            DataTable tableOrders = new DataTable("Orders");
            tableOrders.Columns.Add("CustomerID");
            tableOrders.Columns.Add("ItemName");
            tableOrders.Columns.Add("Quantity");
            tableOrders.Rows.Add(new object[] { 1, "Hawaiian", 2 });
            tableOrders.Rows.Add(new object[] { 2, "Pepperoni", 1 });
            tableOrders.Rows.Add(new object[] { 2, "Chicago", 1 });

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(tableCustomers);
            dataSet.Tables.Add(tableOrders);
            dataSet.Relations.Add(tableCustomers.Columns["CustomerID"], tableOrders.Columns["CustomerID"]);

            using (FileStream streamIn = new FileStream(MyDir + "Mail merge.doc", FileMode.Open, FileAccess.Read))
            {
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MailMergeStreamWithRegionsDataTable.1.docx", FileMode.Create, FileAccess.ReadWrite))
                    MailMerger.ExecuteWithRegions(streamIn, streamOut, SaveFormat.Docx, dataSet);

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MailMergeStreamWithRegionsDataTable.2.docx", FileMode.Create, FileAccess.ReadWrite))
                {
                    MailMergeOptions mailMergeOptions = new MailMergeOptions();
                    mailMergeOptions.TrimWhitespaces = true;
                    MailMerger.ExecuteWithRegions(streamIn, streamOut, SaveFormat.Docx, dataSet, mailMergeOptions);
                }
            }
            //ExEnd:MailMergeStreamWithRegionsDataSet
        }

        [Test]
        public void MailMergeContextStreamWithRegionsDataSet()
        {
            //ExStart:MailMergeContextStreamWithRegionsDataSet
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:MailMerger.Create(MailMergerContext)
            //ExFor:MailMergerContext
            //ExFor:MailMergerContext.SetRegionsDataSource(DataSet)
            //ExSummary:Shows how to do mail merge with regions operation from a DataSet using documents from the stream using context.
            // There is a several ways to do mail merge with regions operation from a DataSet using documents from the stream:
            DataTable tableCustomers = new DataTable("Customers");
            tableCustomers.Columns.Add("CustomerID");
            tableCustomers.Columns.Add("CustomerName");
            tableCustomers.Rows.Add(new object[] { 1, "John Doe" });
            tableCustomers.Rows.Add(new object[] { 2, "Jane Doe" });

            DataTable tableOrders = new DataTable("Orders");
            tableOrders.Columns.Add("CustomerID");
            tableOrders.Columns.Add("ItemName");
            tableOrders.Columns.Add("Quantity");
            tableOrders.Rows.Add(new object[] { 1, "Hawaiian", 2 });
            tableOrders.Rows.Add(new object[] { 2, "Pepperoni", 1 });
            tableOrders.Rows.Add(new object[] { 2, "Chicago", 1 });

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(tableCustomers);
            dataSet.Tables.Add(tableOrders);
            dataSet.Relations.Add(tableCustomers.Columns["CustomerID"], tableOrders.Columns["CustomerID"]);

            using (FileStream streamIn = new FileStream(MyDir + "Mail merge.doc", FileMode.Open, FileAccess.Read))
            {
                MailMergerContext mailMergerContext = new MailMergerContext();
                mailMergerContext.SetRegionsDataSource(dataSet);
                mailMergerContext.MailMergeOptions.TrimWhitespaces = true;

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MailMergeContextStreamWithRegionsDataSet.docx", FileMode.Create, FileAccess.ReadWrite))
                    MailMerger.Create(mailMergerContext)
                    .From(streamIn)
                    .To(streamOut, SaveFormat.Docx)
                    .Execute();
            }
            //ExEnd:MailMergeContextStreamWithRegionsDataSet
        }

        [Test]
        public void MailMergeStreamWithRegionsToImagesDataSet()
        {
            //ExStart:MailMergeStreamWithRegionsToImagesDataSet
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:MailMerger.ExecuteWithRegionsToImages(Stream, ImageSaveOptions, DataSet, MailMergeOptions)
            //ExSummary:Shows how to do mail merge with regions operation from a DataSet using documents from the stream and save result to images.
            // There is a several ways to do mail merge with regions operation from a DataSet using documents from the stream:
            DataTable tableCustomers = new DataTable("Customers");
            tableCustomers.Columns.Add("CustomerID");
            tableCustomers.Columns.Add("CustomerName");
            tableCustomers.Rows.Add(new object[] { 1, "John Doe" });
            tableCustomers.Rows.Add(new object[] { 2, "Jane Doe" });

            DataTable tableOrders = new DataTable("Orders");
            tableOrders.Columns.Add("CustomerID");
            tableOrders.Columns.Add("ItemName");
            tableOrders.Columns.Add("Quantity");
            tableOrders.Rows.Add(new object[] { 1, "Hawaiian", 2 });
            tableOrders.Rows.Add(new object[] { 2, "Pepperoni", 1 });
            tableOrders.Rows.Add(new object[] { 2, "Chicago", 1 });

            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(tableCustomers);
            dataSet.Tables.Add(tableOrders);
            dataSet.Relations.Add(tableCustomers.Columns["CustomerID"], tableOrders.Columns["CustomerID"]);

            using (FileStream streamIn = new FileStream(MyDir + "Mail merge.doc", FileMode.Open, FileAccess.Read))
            {
                Stream[] images = MailMerger.ExecuteWithRegionsToImages(streamIn, new ImageSaveOptions(SaveFormat.Png), dataSet);
                MailMergeOptions mailMergeOptions = new MailMergeOptions();
                mailMergeOptions.TrimWhitespaces = true;
                images = MailMerger.ExecuteWithRegionsToImages(streamIn, new ImageSaveOptions(SaveFormat.Png), dataSet, mailMergeOptions);
            }
            //ExEnd:MailMergeStreamWithRegionsToImagesDataSet
        }

        [Test]
        public void Replace()
        {
            //ExStart:Replace
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Replacer.Replace(String, String, String, String)
            //ExFor:Replacer.Replace(String, String, SaveFormat, String, String, FindReplaceOptions)
            //ExSummary:Shows how to replace string in the document.
            // There is a several ways to replace string in the document:
            string doc = MyDir + "Footer.docx";
            string pattern = "(C)2006 Aspose Pty Ltd.";
            string replacement = "Copyright (C) 2024 by Aspose Pty Ltd.";

            FindReplaceOptions options = new FindReplaceOptions();
            options.FindWholeWordsOnly = false;
            Replacer.Replace(doc, ArtifactsDir + "LowCode.Replace.1.docx", pattern, replacement);
            Replacer.Replace(doc, ArtifactsDir + "LowCode.Replace.2.docx", SaveFormat.Docx, pattern, replacement);
            Replacer.Replace(doc, ArtifactsDir + "LowCode.Replace.3.docx", SaveFormat.Docx, pattern, replacement, options);
            //ExEnd:Replace
        }

        [Test]
        public void ReplaceContext()
        {
            //ExStart:ReplaceContext
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Replacer.Create(ReplacerContext)
            //ExFor:ReplacerContext
            //ExFor:ReplacerContext.SetReplacement(String, String)
            //ExFor:ReplacerContext.FindReplaceOptions
            //ExSummary:Shows how to replace string in the document using context.
            // There is a several ways to replace string in the document:
            string doc = MyDir + "Footer.docx";
            string pattern = "(C)2006 Aspose Pty Ltd.";
            string replacement = "Copyright (C) 2024 by Aspose Pty Ltd.";

            ReplacerContext replacerContext = new ReplacerContext();
            replacerContext.SetReplacement(pattern, replacement);
            replacerContext.FindReplaceOptions.FindWholeWordsOnly = false;

            Replacer.Create(replacerContext)
                .From(doc)
                .To(ArtifactsDir + "LowCode.ReplaceContext.docx")
                .Execute();
            //ExEnd:ReplaceContext
        }

        [Test]
        public void ReplaceToImages()
        {
            //ExStart:ReplaceToImages
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Replacer.ReplaceToImages(String, ImageSaveOptions, String, String, FindReplaceOptions)
            //ExSummary:Shows how to replace string in the document and save result to images.
            // There is a several ways to replace string in the document:
            string doc = MyDir + "Footer.docx";
            string pattern = "(C)2006 Aspose Pty Ltd.";
            string replacement = "Copyright (C) 2024 by Aspose Pty Ltd.";

            Stream[] images = Replacer.ReplaceToImages(doc, new ImageSaveOptions(SaveFormat.Png), pattern, replacement);

            FindReplaceOptions options = new FindReplaceOptions();
            options.FindWholeWordsOnly = false;
            images = Replacer.ReplaceToImages(doc, new ImageSaveOptions(SaveFormat.Png), pattern, replacement, options);
            //ExEnd:ReplaceToImages
        }

        [Test]
        public void ReplaceStream()
        {
            //ExStart:ReplaceStream
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Replacer.Replace(Stream, Stream, SaveFormat, String, String, FindReplaceOptions)
            //ExSummary:Shows how to replace string in the document using documents from the stream.
            // There is a several ways to replace string in the document using documents from the stream:
            string pattern = "(C)2006 Aspose Pty Ltd.";
            string replacement = "Copyright (C) 2024 by Aspose Pty Ltd.";

            using (FileStream streamIn = new FileStream(MyDir + "Footer.docx", FileMode.Open, FileAccess.Read))
            {
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.ReplaceStream.1.docx", FileMode.Create, FileAccess.ReadWrite))
                    Replacer.Replace(streamIn, streamOut, SaveFormat.Docx, pattern, replacement);

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.ReplaceStream.2.docx", FileMode.Create, FileAccess.ReadWrite))
                {
                    FindReplaceOptions options = new FindReplaceOptions();
                    options.FindWholeWordsOnly = false;
                    Replacer.Replace(streamIn, streamOut, SaveFormat.Docx, pattern, replacement, options);
                }
            }
            //ExEnd:ReplaceStream
        }

        [Test]
        public void ReplaceContextStream()
        {
            //ExStart:ReplaceContextStream
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Replacer.Create(ReplacerContext)
            //ExFor:ReplacerContext
            //ExFor:ReplacerContext.SetReplacement(String, String)
            //ExFor:ReplacerContext.FindReplaceOptions
            //ExSummary:Shows how to replace string in the document using documents from the stream using context.
            // There is a several ways to replace string in the document using documents from the stream:
            string pattern = "(C)2006 Aspose Pty Ltd.";
            string replacement = "Copyright (C) 2024 by Aspose Pty Ltd.";

            using (FileStream streamIn = new FileStream(MyDir + "Footer.docx", FileMode.Open, FileAccess.Read))
            {
                ReplacerContext replacerContext = new ReplacerContext();
                replacerContext.SetReplacement(pattern, replacement);
                replacerContext.FindReplaceOptions.FindWholeWordsOnly = false;

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.ReplaceContextStream.docx", FileMode.Create, FileAccess.ReadWrite))
                    Replacer.Create(replacerContext)
                    .From(streamIn)
                    .To(streamOut, SaveFormat.Docx)
                    .Execute();
            }
            //ExEnd:ReplaceContextStream
        }

        [Test]
        public void ReplaceToImagesStream()
        {
            //ExStart:ReplaceToImagesStream
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Replacer.ReplaceToImages(Stream, ImageSaveOptions, String, String, FindReplaceOptions)
            //ExSummary:Shows how to replace string in the document using documents from the stream and save result to images.
            // There is a several ways to replace string in the document using documents from the stream:
            string pattern = "(C)2006 Aspose Pty Ltd.";
            string replacement = "Copyright (C) 2024 by Aspose Pty Ltd.";

            using (FileStream streamIn = new FileStream(MyDir + "Footer.docx", FileMode.Open, FileAccess.Read))
            {
                Stream[] images = Replacer.ReplaceToImages(streamIn, new ImageSaveOptions(SaveFormat.Png), pattern, replacement);

                FindReplaceOptions options = new FindReplaceOptions();
                options.FindWholeWordsOnly = false;
                images = Replacer.ReplaceToImages(streamIn, new ImageSaveOptions(SaveFormat.Png), pattern, replacement, options);
            }
            //ExEnd:ReplaceToImagesStream
        }

        [Test]
        public void ReplaceRegex()
        {
            //ExStart:ReplaceRegex
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Replacer.Replace(String, String, Regex, String)
            //ExFor:Replacer.Replace(String, String, SaveFormat, Regex, String, FindReplaceOptions)
            //ExSummary:Shows how to replace string with regex in the document.
            // There is a several ways to replace string with regex in the document:
            string doc = MyDir + "Footer.docx";
            Regex pattern = new Regex("gr(a|e)y");
            string replacement = "lavender";

            Replacer.Replace(doc, ArtifactsDir + "LowCode.ReplaceRegex.1.docx", pattern, replacement);
            Replacer.Replace(doc, ArtifactsDir + "LowCode.ReplaceRegex.2.docx", SaveFormat.Docx, pattern, replacement);
            FindReplaceOptions options = new FindReplaceOptions();
            options.FindWholeWordsOnly = false;
            Replacer.Replace(doc, ArtifactsDir + "LowCode.ReplaceRegex.3.docx", SaveFormat.Docx, pattern, replacement, options);
            //ExEnd:ReplaceRegex
        }

        [Test]
        public void ReplaceContextRegex()
        {
            //ExStart:ReplaceContextRegex
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Replacer.Create(ReplacerContext)
            //ExFor:ReplacerContext
            //ExFor:ReplacerContext.SetReplacement(Regex, String)
            //ExFor:ReplacerContext.FindReplaceOptions
            //ExSummary:Shows how to replace string with regex in the document using context.
            // There is a several ways to replace string with regex in the document:
            string doc = MyDir + "Footer.docx";
            Regex pattern = new Regex("gr(a|e)y");
            string replacement = "lavender";

            ReplacerContext replacerContext = new ReplacerContext();
            replacerContext.SetReplacement(pattern, replacement);
            replacerContext.FindReplaceOptions.FindWholeWordsOnly = false;

            Replacer.Create(replacerContext)
                .From(doc)
                .To(ArtifactsDir + "LowCode.ReplaceContextRegex.docx")
                .Execute();
            //ExEnd:ReplaceContextRegex
        }

        [Test]
        public void ReplaceToImagesRegex()
        {
            //ExStart:ReplaceToImagesRegex
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Replacer.ReplaceToImages(String, ImageSaveOptions, Regex, String, FindReplaceOptions)
            //ExSummary:Shows how to replace string with regex in the document and save result to images.
            // There is a several ways to replace string with regex in the document:
            string doc = MyDir + "Footer.docx";
            Regex pattern = new Regex("gr(a|e)y");
            string replacement = "lavender";

            Stream[] images = Replacer.ReplaceToImages(doc, new ImageSaveOptions(SaveFormat.Png), pattern, replacement);
            FindReplaceOptions options = new FindReplaceOptions();
            options.FindWholeWordsOnly = false;
            images = Replacer.ReplaceToImages(doc, new ImageSaveOptions(SaveFormat.Png), pattern, replacement, options);
            //ExEnd:ReplaceToImagesRegex
        }

        [Test]
        public void ReplaceStreamRegex()
        {
            //ExStart:ReplaceStreamRegex
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Replacer.Replace(Stream, Stream, SaveFormat, Regex, String, FindReplaceOptions)
            //ExSummary:Shows how to replace string with regex in the document using documents from the stream.
            // There is a several ways to replace string with regex in the document using documents from the stream:
            Regex pattern = new Regex("gr(a|e)y");
            string replacement = "lavender";

            using (FileStream streamIn = new FileStream(MyDir + "Replace regex.docx", FileMode.Open, FileAccess.Read))
            {
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.ReplaceStreamRegex.1.docx", FileMode.Create, FileAccess.ReadWrite))
                    Replacer.Replace(streamIn, streamOut, SaveFormat.Docx, pattern, replacement);

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.ReplaceStreamRegex.2.docx", FileMode.Create, FileAccess.ReadWrite))
                {
                    FindReplaceOptions options = new FindReplaceOptions();
                    options.FindWholeWordsOnly = false;
                    Replacer.Replace(streamIn, streamOut, SaveFormat.Docx, pattern, replacement, options);
                }
            }
            //ExEnd:ReplaceStreamRegex
        }

        [Test]
        public void ReplaceContextStreamRegex()
        {
            //ExStart:ReplaceContextStreamRegex
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Replacer.Create(ReplacerContext)
            //ExFor:ReplacerContext
            //ExFor:ReplacerContext.SetReplacement(Regex, String)
            //ExFor:ReplacerContext.FindReplaceOptions
            //ExSummary:Shows how to replace string with regex in the document using documents from the stream using context.
            // There is a several ways to replace string with regex in the document using documents from the stream:
            Regex pattern = new Regex("gr(a|e)y");
            string replacement = "lavender";

            using (FileStream streamIn = new FileStream(MyDir + "Replace regex.docx", FileMode.Open, FileAccess.Read))
            {
                ReplacerContext replacerContext = new ReplacerContext();
                replacerContext.SetReplacement(pattern, replacement);
                replacerContext.FindReplaceOptions.FindWholeWordsOnly = false;

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.ReplaceContextStreamRegex.docx", FileMode.Create, FileAccess.ReadWrite))
                    Replacer.Create(replacerContext)
                        .From(streamIn)
                        .To(streamOut, SaveFormat.Docx)
                        .Execute();
            }
            //ExEnd:ReplaceContextStreamRegex
        }

        [Test]
        public void ReplaceToImagesStreamRegex()
        {
            //ExStart:ReplaceToImagesStreamRegex
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Replacer.ReplaceToImages(Stream, ImageSaveOptions, Regex, String, FindReplaceOptions)
            //ExSummary:Shows how to replace string with regex in the document using documents from the stream and save result to images.
            // There is a several ways to replace string with regex in the document using documents from the stream:
            Regex pattern = new Regex("gr(a|e)y");
            string replacement = "lavender";

            using (FileStream streamIn = new FileStream(MyDir + "Replace regex.docx", FileMode.Open, FileAccess.Read))
            {
                Stream[] images = Replacer.ReplaceToImages(streamIn, new ImageSaveOptions(SaveFormat.Png), pattern, replacement);
                FindReplaceOptions options = new FindReplaceOptions();
                options.FindWholeWordsOnly = false;
                images = Replacer.ReplaceToImages(streamIn, new ImageSaveOptions(SaveFormat.Png), pattern, replacement, options);
            }
            //ExEnd:ReplaceToImagesStreamRegex
        }

        //ExStart:BuildReportData
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:ReportBuilderOptions
        //ExFor:ReportBuilderOptions.Options
        //ExFor:ReportBuilder.BuildReport(String, String, Object, ReportBuilderOptions)
        //ExFor:ReportBuilder.BuildReport(String, String, SaveFormat, Object, ReportBuilderOptions)
        //ExSummary:Shows how to populate document with data.
        [Test] //ExSkip
        public void BuildReportData()
        {
            // There is a several ways to populate document with data:
            string doc = MyDir + "Reporting engine template - If greedy.docx";

            AsposeData obj = new AsposeData();
            obj.List = new List<string> { "abc" };

            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportWithObject.1.docx", obj);
            ReportBuilderOptions reportBuilderOptions = new ReportBuilderOptions();
            reportBuilderOptions.Options = ReportBuildOptions.AllowMissingMembers;
            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportWithObject.2.docx", obj, reportBuilderOptions);
            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportWithObject.3.docx", SaveFormat.Docx, obj);
            ReportBuilderOptions reportBuilderOptions2 = new ReportBuilderOptions();
            reportBuilderOptions2.Options = ReportBuildOptions.AllowMissingMembers;
            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportWithObject.4.docx", SaveFormat.Docx, obj, reportBuilderOptions2);
        }

        public class AsposeData
        {
            public List<string> List { get; set; }
        }
        //ExEnd:BuildReportData

        [Test]
        public void BuildReportDataStream()
        {
            //ExStart:BuildReportDataStream
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object, ReportBuilderOptions)
            //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object[], String[], ReportBuilderOptions)
            //ExSummary:Shows how to populate document with data using documents from the stream.
            // There is a several ways to populate document with data using documents from the stream:
            AsposeData obj = new AsposeData();
            obj.List = new List<string> { "abc" };

            using (FileStream streamIn = new FileStream(MyDir + "Reporting engine template - If greedy.docx", FileMode.Open, FileAccess.Read))
            {
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.BuildReportDataStream.1.docx", FileMode.Create, FileAccess.ReadWrite))
                    ReportBuilder.BuildReport(streamIn, streamOut, SaveFormat.Docx, obj);

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.BuildReportDataStream.2.docx", FileMode.Create, FileAccess.ReadWrite))
                {
                    ReportBuilderOptions reportBuilderOptions = new ReportBuilderOptions();
                    reportBuilderOptions.Options = ReportBuildOptions.AllowMissingMembers;
                    ReportBuilder.BuildReport(streamIn, streamOut, SaveFormat.Docx, obj, reportBuilderOptions);
                }

                MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "Hello World");
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.BuildReportDataStream.3.docx", FileMode.Create, FileAccess.ReadWrite))
                {
                    ReportBuilderOptions reportBuilderOptions2 = new ReportBuilderOptions();
                    reportBuilderOptions2.Options = ReportBuildOptions.AllowMissingMembers;
                    ReportBuilder.BuildReport(streamIn, streamOut, SaveFormat.Docx, new object[] { sender }, new[] { "s" }, reportBuilderOptions2);
                }
            }
            //ExEnd:BuildReportDataStream
        }

        //ExStart:BuildReportDataSource
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:ReportBuilder.BuildReport(String, String, Object, String, ReportBuilderOptions)
        //ExFor:ReportBuilder.BuildReport(String, String, SaveFormat, Object, String, ReportBuilderOptions)
        //ExFor:ReportBuilder.BuildReport(String, String, Object[], String[], ReportBuilderOptions)
        //ExFor:ReportBuilder.BuildReport(String, String, SaveFormat, Object[], String[], ReportBuilderOptions)
        //ExFor:ReportBuilder.BuildReportToImages(String, ImageSaveOptions, Object[], String[], ReportBuilderOptions)
        //ExFor:ReportBuilder.Create(ReportBuilderContext)
        //ExFor:ReportBuilderContext
        //ExFor:ReportBuilderContext.ReportBuilderOptions
        //ExFor:ReportBuilderContext.DataSources
        //ExSummary:Shows how to populate document with data sources.
        [Test] //ExSkip
        public void BuildReportDataSource()
        {
            // There is a several ways to populate document with data sources:
            string doc = MyDir + "Report building.docx";

            MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "Hello World");

            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportDataSource.1.docx", sender, "s");
            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportDataSource.2.docx", new object[] { sender }, new[] { "s" });
            ReportBuilderOptions reportBuilderOptions = new ReportBuilderOptions();
            reportBuilderOptions.Options = ReportBuildOptions.AllowMissingMembers;
            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportDataSource.3.docx", sender, "s", reportBuilderOptions);
            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportDataSource.4.docx", SaveFormat.Docx, sender, "s");
            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportDataSource.5.docx", SaveFormat.Docx, new object[] { sender }, new[] { "s" });
            ReportBuilderOptions reportBuilderOptions2 = new ReportBuilderOptions();
            reportBuilderOptions2.Options = ReportBuildOptions.AllowMissingMembers;
            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportDataSource.6.docx", SaveFormat.Docx, sender, "s", reportBuilderOptions2);
            ReportBuilderOptions reportBuilderOptions3 = new ReportBuilderOptions();
            reportBuilderOptions3.Options = ReportBuildOptions.AllowMissingMembers;
            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportDataSource.7.docx", SaveFormat.Docx, new object[] { sender }, new[] { "s" }, reportBuilderOptions3);
            ReportBuilderOptions reportBuilderOptions4 = new ReportBuilderOptions();
            reportBuilderOptions4.Options = ReportBuildOptions.AllowMissingMembers;
            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportDataSource.8.docx", new object[] { sender }, new[] { "s" }, reportBuilderOptions4);
            ReportBuilderOptions reportBuilderOptions5 = new ReportBuilderOptions();
            reportBuilderOptions5.Options = ReportBuildOptions.AllowMissingMembers;

            Stream[] images = ReportBuilder.BuildReportToImages(doc, new ImageSaveOptions(SaveFormat.Png), new object[] { sender }, new[] { "s" }, reportBuilderOptions5);

            ReportBuilderContext reportBuilderContext = new ReportBuilderContext();
            reportBuilderContext.ReportBuilderOptions.MissingMemberMessage = "Missed members";
            reportBuilderContext.DataSources.Add(sender, "s");

            ReportBuilder.Create(reportBuilderContext)
                .From(doc)
                .To(ArtifactsDir + "LowCode.BuildReportDataSource.9.docx")
                .Execute();
        }

        public class MessageTestClass
        {
            public string Name { get; set; }
            public string Message { get; set; }

            public MessageTestClass(string name, string message)
            {
                Name = name;
                Message = message;
            }
        }
        //ExEnd:BuildReportDataSource

        [Test]
        public void BuildReportDataSourceStream()
        {
            //ExStart:BuildReportDataSourceStream
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object, String, ReportBuilderOptions)
            //ExFor:ReportBuilder.BuildReportToImages(Stream, ImageSaveOptions, Object[], String[], ReportBuilderOptions)
            //ExFor:ReportBuilder.Create(ReportBuilderContext)
            //ExFor:ReportBuilderContext
            //ExFor:ReportBuilderContext.ReportBuilderOptions
            //ExFor:ReportBuilderContext.DataSources
            //ExSummary:Shows how to populate document with data sources using documents from the stream.
            // There is a several ways to populate document with data sources using documents from the stream:
            MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "Hello World");

            using (FileStream streamIn = new FileStream(MyDir + "Report building.docx", FileMode.Open, FileAccess.Read))
            {
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.BuildReportDataSourceStream.1.docx", FileMode.Create, FileAccess.ReadWrite))
                    ReportBuilder.BuildReport(streamIn, streamOut, SaveFormat.Docx, new object[] { sender }, new[] { "s" });

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.BuildReportDataSourceStream.2.docx", FileMode.Create, FileAccess.ReadWrite))
                    ReportBuilder.BuildReport(streamIn, streamOut, SaveFormat.Docx, sender, "s");

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.BuildReportDataSourceStream.3.docx", FileMode.Create, FileAccess.ReadWrite))
                {
                    ReportBuilderOptions reportBuilderOptions = new ReportBuilderOptions();
                    reportBuilderOptions.Options = ReportBuildOptions.AllowMissingMembers;
                    ReportBuilder.BuildReport(streamIn, streamOut, SaveFormat.Docx, sender, "s", reportBuilderOptions);
                }
                ReportBuilderOptions reportBuilderOptions2 = new ReportBuilderOptions();
                reportBuilderOptions2.Options = ReportBuildOptions.AllowMissingMembers;

                Stream[] images = ReportBuilder.BuildReportToImages(streamIn, new ImageSaveOptions(SaveFormat.Png), new object[] { sender }, new[] { "s" }, reportBuilderOptions2);

                ReportBuilderContext reportBuilderContext = new ReportBuilderContext();
                reportBuilderContext.ReportBuilderOptions.MissingMemberMessage = "Missed members";
                reportBuilderContext.DataSources.Add(sender, "s");

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.BuildReportDataSourceStream.4.docx", FileMode.Create, FileAccess.ReadWrite))
                    ReportBuilder.Create(reportBuilderContext)
                        .From(streamIn)
                        .To(streamOut, SaveFormat.Docx)
                        .Execute();
            }
            //ExEnd:BuildReportDataSourceStream
        }

        [Test]
        public void RemoveBlankPages()
        {
            //ExStart:RemoveBlankPages
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Splitter.RemoveBlankPages(String, String)
            //ExFor:Splitter.RemoveBlankPages(String, String, SaveFormat)
            //ExSummary:Shows how to remove empty pages from the document.
            // There is a several ways to remove empty pages from the document:
            string doc = MyDir + "Blank pages.docx";

            Splitter.RemoveBlankPages(doc, ArtifactsDir + "LowCode.RemoveBlankPages.1.docx");
            Splitter.RemoveBlankPages(doc, ArtifactsDir + "LowCode.RemoveBlankPages.2.docx", SaveFormat.Docx);
            //ExEnd:RemoveBlankPages
        }

        [Test]
        public void RemoveBlankPagesStream()
        {
            //ExStart:RemoveBlankPagesStream
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Splitter.RemoveBlankPages(Stream, Stream, SaveFormat)
            //ExSummary:Shows how to remove empty pages from the document from the stream.
            using (FileStream streamIn = new FileStream(MyDir + "Blank pages.docx", FileMode.Open, FileAccess.Read))
            {
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.RemoveBlankPagesStream.docx", FileMode.Create, FileAccess.ReadWrite))
                    Splitter.RemoveBlankPages(streamIn, streamOut, SaveFormat.Docx);
            }
            //ExEnd:RemoveBlankPagesStream
        }

        [Test]
        public void ExtractPages()
        {
            //ExStart:ExtractPages
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Splitter.ExtractPages(String, String, int, int)
            //ExFor:Splitter.ExtractPages(String, String, SaveFormat, int, int)
            //ExSummary:Shows how to extract pages from the document.
            // There is a several ways to extract pages from the document:
            string doc = MyDir + "Big document.docx";

            Splitter.ExtractPages(doc, ArtifactsDir + "LowCode.ExtractPages.1.docx", 0, 2);
            Splitter.ExtractPages(doc, ArtifactsDir + "LowCode.ExtractPages.2.docx", SaveFormat.Docx, 0, 2);
            //ExEnd:ExtractPages
        }

        [Test]
        public void ExtractPagesStream()
        {
            //ExStart:ExtractPagesStream
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Splitter.ExtractPages(Stream, Stream, SaveFormat, int, int)
            //ExSummary:Shows how to extract pages from the document from the stream.
            using (FileStream streamIn = new FileStream(MyDir + "Big document.docx", FileMode.Open, FileAccess.Read))
            {
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.ExtractPagesStream.docx", FileMode.Create, FileAccess.ReadWrite))
                    Splitter.ExtractPages(streamIn, streamOut, SaveFormat.Docx, 0, 2);
            }
            //ExEnd:ExtractPagesStream
        }

        [Test]
        public void SplitDocument()
        {
            //ExStart:SplitDocument
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:SplitCriteria
            //ExFor:SplitOptions.SplitCriteria
            //ExFor:Splitter.Split(String, String, SplitOptions)
            //ExFor:Splitter.Split(String, String, SaveFormat, SplitOptions)
            //ExSummary:Shows how to split document by pages.
            string doc = MyDir + "Big document.docx";

            SplitOptions options = new SplitOptions();
            options.SplitCriteria = SplitCriteria.Page;
            Splitter.Split(doc, ArtifactsDir + "LowCode.SplitDocument.1.docx", options);
            Splitter.Split(doc, ArtifactsDir + "LowCode.SplitDocument.2.docx", SaveFormat.Docx, options);
            //ExEnd:SplitDocument
        }

        [Test]
        public void SplitContextDocument()
        {
            //ExStart:SplitContextDocument
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Splitter.Create(SplitterContext)
            //ExFor:SplitterContext
            //ExFor:SplitterContext.SplitOptions
            //ExSummary:Shows how to split document by pages using context.
            string doc = MyDir + "Big document.docx";

            SplitterContext splitterContext = new SplitterContext();
            splitterContext.SplitOptions.SplitCriteria = SplitCriteria.Page;

            Splitter.Create(splitterContext)
                .From(doc)
                .To(ArtifactsDir + "LowCode.SplitContextDocument.docx")
                .Execute();
            //ExEnd:SplitContextDocument
        }

        [Test]
        public void SplitDocumentStream()
        {
            //ExStart:SplitDocumentStream
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Splitter.Split(Stream, SaveFormat, SplitOptions)
            //ExSummary:Shows how to split document from the stream by pages.
            using (FileStream streamIn = new FileStream(MyDir + "Big document.docx", FileMode.Open, FileAccess.Read))
            {
                SplitOptions options = new SplitOptions();
                options.SplitCriteria = SplitCriteria.Page;
                Stream[] stream = Splitter.Split(streamIn, SaveFormat.Docx, options);
            }
            //ExEnd:SplitDocumentStream
        }

        [Test]
        public void SplitContextDocumentStream()
        {
            //ExStart:SplitContextDocumentStream
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Splitter.Create(SplitterContext)
            //ExFor:SplitterContext
            //ExFor:SplitterContext.SplitOptions
            //ExSummary:Shows how to split document from the stream by pages using context.
            using (FileStream streamIn = new FileStream(MyDir + "Big document.docx", FileMode.Open, FileAccess.Read))
            {
                SplitterContext splitterContext = new SplitterContext();
                splitterContext.SplitOptions.SplitCriteria = SplitCriteria.Page;

                List<Stream> pages = new List<Stream>();
                Splitter.Create(splitterContext)
                    .From(streamIn)
                    .To(pages, SaveFormat.Docx)
                    .Execute();
            }
            //ExEnd:SplitContextDocumentStream
        }

        [Test]
        public void WatermarkText()
        {
            //ExStart:WatermarkText
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Watermarker.SetText(String, String, String)
            //ExFor:Watermarker.SetText(String, String, String, TextWatermarkOptions)
            //ExFor:Watermarker.SetText(String, String, SaveFormat, String, TextWatermarkOptions)
            //ExSummary:Shows how to insert watermark text to the document.
            string doc = MyDir + "Big document.docx";
            string watermarkText = "This is a watermark";

            Watermarker.SetText(doc, ArtifactsDir + "LowCode.WatermarkText.1.docx", watermarkText);
            Watermarker.SetText(doc, ArtifactsDir + "LowCode.WatermarkText.2.docx", SaveFormat.Docx, watermarkText);
            TextWatermarkOptions watermarkOptions = new TextWatermarkOptions();
            watermarkOptions.Color = Color.Red;
            Watermarker.SetText(doc, ArtifactsDir + "LowCode.WatermarkText.3.docx", watermarkText, watermarkOptions);
            Watermarker.SetText(doc, ArtifactsDir + "LowCode.WatermarkText.4.docx", SaveFormat.Docx, watermarkText, watermarkOptions);
            //ExEnd:WatermarkText
        }

        [Test]
        public void WatermarkContextText()
        {
            //ExStart:WatermarkContextText
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Watermarker.Create(WatermarkerContext)
            //ExFor:WatermarkerContext
            //ExFor:WatermarkerContext.TextWatermark
            //ExFor:WatermarkerContext.TextWatermarkOptions
            //ExSummary:Shows how to insert watermark text to the document using context.
            string doc = MyDir + "Big document.docx";
            string watermarkText = "This is a watermark";

            WatermarkerContext watermarkerContext = new WatermarkerContext();
            watermarkerContext.TextWatermark = watermarkText;

            watermarkerContext.TextWatermarkOptions.Color = Color.Red;

            Watermarker.Create(watermarkerContext)
                .From(doc)
                .To(ArtifactsDir + "LowCode.WatermarkContextText.docx")
                .Execute();
            //ExEnd:WatermarkContextText
        }

        [Test]
        public void WatermarkTextStream()
        {
            //ExStart:WatermarkTextStream
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Watermarker.SetText(Stream, Stream, SaveFormat, String, TextWatermarkOptions)
            //ExSummary:Shows how to insert watermark text to the document from the stream.
            string watermarkText = "This is a watermark";

            using (FileStream streamIn = new FileStream(MyDir + "Document.docx", FileMode.Open, FileAccess.Read))
            {
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.WatermarkTextStream.1.docx", FileMode.Create, FileAccess.ReadWrite))
                    Watermarker.SetText(streamIn, streamOut, SaveFormat.Docx, watermarkText);

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.WatermarkTextStream.2.docx", FileMode.Create, FileAccess.ReadWrite))
                {
                    TextWatermarkOptions options = new TextWatermarkOptions();
                    options.Color = Color.Red;
                    Watermarker.SetText(streamIn, streamOut, SaveFormat.Docx, watermarkText, options);
                }
            }
            //ExEnd:WatermarkTextStream
        }

        [Test]
        public void WatermarkContextTextStream()
        {
            //ExStart:WatermarkContextTextStream
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Watermarker.Create(WatermarkerContext)
            //ExFor:WatermarkerContext
            //ExFor:WatermarkerContext.TextWatermark
            //ExFor:WatermarkerContext.TextWatermarkOptions
            //ExSummary:Shows how to insert watermark text to the document from the stream using context.
            string watermarkText = "This is a watermark";

            using (FileStream streamIn = new FileStream(MyDir + "Document.docx", FileMode.Open, FileAccess.Read))
            {
                WatermarkerContext watermarkerContext = new WatermarkerContext();
                watermarkerContext.TextWatermark = watermarkText;

                watermarkerContext.TextWatermarkOptions.Color = Color.Red;

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.WatermarkContextTextStream.docx", FileMode.Create, FileAccess.ReadWrite))
                    Watermarker.Create(watermarkerContext)
                        .From(streamIn)
                        .To(streamOut, SaveFormat.Docx)
                        .Execute();
            }
            //ExEnd:WatermarkContextTextStream
        }

        [Test]
        public void WatermarkImage()
        {
            //ExStart:WatermarkImage
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Watermarker.SetImage(String, String, String)
            //ExFor:Watermarker.SetImage(String, String, String, ImageWatermarkOptions)
            //ExFor:Watermarker.SetImage(String, String, SaveFormat, String, ImageWatermarkOptions)
            //ExSummary:Shows how to insert watermark image to the document.
            string doc = MyDir + "Document.docx";
            string watermarkImage = ImageDir + "Logo.jpg";

            Watermarker.SetImage(doc, ArtifactsDir + "LowCode.SetWatermarkImage.1.docx", watermarkImage);
            Watermarker.SetImage(doc, ArtifactsDir + "LowCode.SetWatermarkText.2.docx", SaveFormat.Docx, watermarkImage);

            ImageWatermarkOptions options = new ImageWatermarkOptions();
            options.Scale = 50;
            Watermarker.SetImage(doc, ArtifactsDir + "LowCode.SetWatermarkText.3.docx", watermarkImage, options);
            Watermarker.SetImage(doc, ArtifactsDir + "LowCode.SetWatermarkText.4.docx", SaveFormat.Docx, watermarkImage, options);
            //ExEnd:WatermarkImage
        }

        [Test]
        public void WatermarkContextImage()
        {
            //ExStart:WatermarkContextImage
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Watermarker.Create(WatermarkerContext)
            //ExFor:WatermarkerContext
            //ExFor:WatermarkerContext.ImageWatermark
            //ExFor:WatermarkerContext.ImageWatermarkOptions
            //ExSummary:Shows how to insert watermark image to the document using context.
            string doc = MyDir + "Document.docx";
            string watermarkImage = ImageDir + "Logo.jpg";


            WatermarkerContext watermarkerContext = new WatermarkerContext();
            watermarkerContext.ImageWatermark = File.ReadAllBytes(watermarkImage);

            watermarkerContext.ImageWatermarkOptions.Scale = 50;

            Watermarker.Create(watermarkerContext)
                .From(doc)
                .To(ArtifactsDir + "LowCode.WatermarkContextImage.docx")
                .Execute();
            //ExEnd:WatermarkContextImage
        }

        [Test]
        public void WatermarkImageStream()
        {
            //ExStart:WatermarkImageStream
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Watermarker.SetImage(Stream, Stream, SaveFormat, Image, ImageWatermarkOptions)
            //ExSummary:Shows how to insert watermark image to the document from a stream.
            using (FileStream streamIn = new FileStream(MyDir + "Document.docx", FileMode.Open, FileAccess.Read))
            {
#if NET461_OR_GREATER || JAVA //ExSkip
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.SetWatermarkText.1.docx", FileMode.Create, FileAccess.ReadWrite))
                    Watermarker.SetImage(streamIn, streamOut, SaveFormat.Docx, System.Drawing.Image.FromFile(ImageDir + "Logo.jpg"));
#endif //ExSkip

#if NET461_OR_GREATER || JAVA //ExSkip
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.SetWatermarkText.2.docx", FileMode.Create, FileAccess.ReadWrite))
                                                      {
                    ImageWatermarkOptions options = new ImageWatermarkOptions();
                    options.Scale = 50;
                    Watermarker.SetImage(streamIn, streamOut, SaveFormat.Docx, System.Drawing.Image.FromFile(ImageDir + "Logo.jpg"), options);
                                                      }
#endif //ExSkip
            }
            //ExEnd:WatermarkImageStream
        }

        [Test]
        public void WatermarkContextImageStream()
        {
            //ExStart:WatermarkContextImageStream
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Watermarker.Create(WatermarkerContext)
            //ExFor:WatermarkerContext
            //ExFor:WatermarkerContext.ImageWatermark
            //ExFor:WatermarkerContext.ImageWatermarkOptions
            //ExSummary:Shows how to insert watermark image to the document from a stream using context.
            string watermarkImage = ImageDir + "Logo.jpg";

            using (FileStream streamIn = new FileStream(MyDir + "Document.docx", FileMode.Open, FileAccess.Read))
            {
                WatermarkerContext watermarkerContext = new WatermarkerContext();
                watermarkerContext.ImageWatermark = File.ReadAllBytes(watermarkImage);

                watermarkerContext.ImageWatermarkOptions.Scale = 50;

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.WatermarkContextImageStream.docx", FileMode.Create, FileAccess.ReadWrite))
                    Watermarker.Create(watermarkerContext)
                        .From(streamIn)
                        .To(streamOut, SaveFormat.Docx)
                        .Execute();
            }
            //ExEnd:WatermarkContextImageStream
        }

        [Test]
        public void WatermarkTextToImages()
        {
            //ExStart:WatermarkTextToImages
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Watermarker.SetWatermarkToImages(String, ImageSaveOptions, String, TextWatermarkOptions)
            //ExSummary:Shows how to insert watermark text to the document and save result to images.
            string doc = MyDir + "Big document.docx";
            string watermarkText = "This is a watermark";

            Stream[] images = Watermarker.SetWatermarkToImages(doc, new ImageSaveOptions(SaveFormat.Png), watermarkText);

            TextWatermarkOptions watermarkOptions = new TextWatermarkOptions();
            watermarkOptions.Color = Color.Red;
            images = Watermarker.SetWatermarkToImages(doc, new ImageSaveOptions(SaveFormat.Png), watermarkText, watermarkOptions);
            //ExEnd:WatermarkTextToImages
        }

        [Test]
        public void WatermarkTextToImagesStream()
        {
            //ExStart:WatermarkTextToImagesStream
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Watermarker.SetWatermarkToImages(Stream, ImageSaveOptions, String, TextWatermarkOptions)
            //ExSummary:Shows how to insert watermark text to the document from the stream and save result to images.
            string watermarkText = "This is a watermark";

            using (FileStream streamIn = new FileStream(MyDir + "Document.docx", FileMode.Open, FileAccess.Read))
            {
                Stream[] images = Watermarker.SetWatermarkToImages(streamIn, new ImageSaveOptions(SaveFormat.Png), watermarkText);

                TextWatermarkOptions watermarkOptions = new TextWatermarkOptions();
                watermarkOptions.Color = Color.Red;
                images = Watermarker.SetWatermarkToImages(streamIn, new ImageSaveOptions(SaveFormat.Png), watermarkText, watermarkOptions);
            }
            //ExEnd:WatermarkTextToImagesStream
        }

        [Test]
        public void WatermarkImageToImages()
        {
            //ExStart:WatermarkImageToImages
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Watermarker.SetWatermarkToImages(String, ImageSaveOptions, Byte[], ImageWatermarkOptions)
            //ExSummary:Shows how to insert watermark image to the document and save result to images.
            string doc = MyDir + "Document.docx";
            string watermarkImage = ImageDir + "Logo.jpg";

            Watermarker.SetWatermarkToImages(doc, new ImageSaveOptions(SaveFormat.Png), File.ReadAllBytes(watermarkImage));

            ImageWatermarkOptions options = new ImageWatermarkOptions();
            options.Scale = 50;
            Watermarker.SetWatermarkToImages(doc, new ImageSaveOptions(SaveFormat.Png), File.ReadAllBytes(watermarkImage), options);
            //ExEnd:WatermarkImageToImages
        }

        [Test]
        public void WatermarkImageToImagesStream()
        {
            //ExStart:WatermarkImageToImagesStream
            //GistId:12a3a3cfe30f3145220db88428a9f814
            //ExFor:Watermarker.SetWatermarkToImages(Stream, ImageSaveOptions, Stream, ImageWatermarkOptions)
            //ExSummary:Shows how to insert watermark image to the document from a stream and save result to images.
            string watermarkImage = ImageDir + "Logo.jpg";

            using (FileStream streamIn = new FileStream(MyDir + "Document.docx", FileMode.Open, FileAccess.Read))
            {
                using (FileStream imageStream = new FileStream(watermarkImage, FileMode.Open, FileAccess.Read))
                {
                    Watermarker.SetWatermarkToImages(streamIn, new ImageSaveOptions(SaveFormat.Png), imageStream);
                    ImageWatermarkOptions options = new ImageWatermarkOptions();
                    options.Scale = 50;
                    Watermarker.SetWatermarkToImages(streamIn, new ImageSaveOptions(SaveFormat.Png), imageStream, options);
                }
            }
            //ExEnd:WatermarkImageToImagesStream
        }
    }
}
