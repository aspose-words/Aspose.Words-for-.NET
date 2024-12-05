// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
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
using Aspose.Words.LowCode.MailMerging;
using Aspose.Words.LowCode.Reporting;
using Aspose.Words.LowCode.Splitting;
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
            //ExFor:Merger.Merge(String, String[], SaveOptions, MergeFormatMode)
            //ExFor:Merger.Merge(String, String[], SaveFormat, MergeFormatMode)
            //ExFor:LowCode.MergeFormatMode
            //ExFor:LowCode.Merger
            //ExSummary:Shows how to merge documents into a single output document.
            //There is a several ways to merge documents:
            Merger.Merge(ArtifactsDir + "LowCode.MergeDocument.SimpleMerge.docx", new[] { MyDir + "Big document.docx", MyDir + "Tables.docx" });

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions { Password = "Aspose.Words" };
            Merger.Merge(ArtifactsDir + "LowCode.MergeDocument.SaveOptions.docx", new[] { MyDir + "Big document.docx", MyDir + "Tables.docx" }, saveOptions, MergeFormatMode.KeepSourceFormatting);

            Merger.Merge(ArtifactsDir + "LowCode.MergeDocument.SaveFormat.pdf", new[] { MyDir + "Big document.docx", MyDir + "Tables.docx" }, SaveFormat.Pdf, MergeFormatMode.KeepSourceLayout);

            Document doc = Merger.Merge(new[] { MyDir + "Big document.docx", MyDir + "Tables.docx" }, MergeFormatMode.MergeFormatting);
            doc.Save(ArtifactsDir + "LowCode.MergeDocument.DocumentInstance.docx");
            //ExEnd
        }

        [Test]
        public void MergeStreamDocument()
        {
            //ExStart
            //ExFor:Merger.Merge(Stream[], MergeFormatMode)
            //ExFor:Merger.Merge(Stream, Stream[], SaveOptions, MergeFormatMode)
            //ExFor:Merger.Merge(Stream, Stream[], SaveFormat)
            //ExSummary:Shows how to merge documents from stream into a single output document.
            //There is a several ways to merge documents from stream:
            using (FileStream firstStreamIn = new FileStream(MyDir + "Big document.docx", FileMode.Open, FileAccess.Read))
            {
                using (FileStream secondStreamIn = new FileStream(MyDir + "Tables.docx", FileMode.Open, FileAccess.Read))
                {
                    OoxmlSaveOptions saveOptions = new OoxmlSaveOptions { Password = "Aspose.Words" };
                    using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MergeStreamDocument.SaveOptions.docx", FileMode.Create, FileAccess.ReadWrite))
                        Merger.Merge(streamOut, new[] { firstStreamIn, secondStreamIn }, saveOptions, MergeFormatMode.KeepSourceFormatting);

                    using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MergeStreamDocument.SaveFormat.docx", FileMode.Create, FileAccess.ReadWrite))
                        Merger.Merge(streamOut, new[] { firstStreamIn, secondStreamIn }, SaveFormat.Docx);

                    Document doc = Merger.Merge(new[] { firstStreamIn, secondStreamIn }, MergeFormatMode.MergeFormatting);
                    doc.Save(ArtifactsDir + "LowCode.MergeStreamDocument.DocumentInstance.docx");
                }
            }
            //ExEnd
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
            //ExSummary:Shows how to convert documents with a single line of code.
            Converter.Convert(MyDir + "Document.docx", ArtifactsDir + "LowCode.Convert.pdf");

            Converter.Convert(MyDir + "Document.docx", ArtifactsDir + "LowCode.Convert.rtf", SaveFormat.Rtf);

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions { Password = "Aspose.Words" };
            Converter.Convert(MyDir + "Document.doc", ArtifactsDir + "LowCode.Convert.docx", saveOptions);
            //ExEnd:Convert
        }

        [Test]
        public void ConvertStream()
        {
            //ExStart:ConvertStream
            //GistId:708ce40a68fac5003d46f6b4acfd5ff1
            //ExFor:Converter.Convert(Stream, Stream, SaveFormat)
            //ExFor:Converter.Convert(Stream, Stream, SaveOptions)
            //ExSummary:Shows how to convert documents with a single line of code (Stream).
            using (FileStream streamIn = new FileStream(MyDir + "Big document.docx", FileMode.Open, FileAccess.Read))
            {
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.ConvertStream.SaveFormat.docx", FileMode.Create, FileAccess.ReadWrite))
                    Converter.Convert(streamIn, streamOut, SaveFormat.Docx);

                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions { Password = "Aspose.Words" };
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.ConvertStream.SaveOptions.docx", FileMode.Create, FileAccess.ReadWrite))
                    Converter.Convert(streamIn, streamOut, saveOptions);
            }
            //ExEnd:ConvertStream
        }

        [Test]
        public void ConvertToImages()
        {
            //ExStart:ConvertToImages
            //GistId:708ce40a68fac5003d46f6b4acfd5ff1
            //ExFor:Converter.ConvertToImages(String, String)
            //ExFor:Converter.ConvertToImages(String, String, SaveFormat)
            //ExFor:Converter.ConvertToImages(String, String, ImageSaveOptions)
            //ExSummary:Shows how to convert document to images.
            Converter.ConvertToImages(MyDir + "Big document.docx", ArtifactsDir + "LowCode.ConvertToImages.png");

            Converter.ConvertToImages(MyDir + "Big document.docx", ArtifactsDir + "LowCode.ConvertToImages.jpeg", SaveFormat.Jpeg);

            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Png);
            imageSaveOptions.PageSet = new PageSet(1);
            Converter.ConvertToImages(MyDir + "Big document.docx", ArtifactsDir + "LowCode.ConvertToImages.png", imageSaveOptions);
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
            Stream[] streams = Converter.ConvertToImages(MyDir + "Big document.docx", SaveFormat.Png);

            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Png);
            imageSaveOptions.PageSet = new PageSet(1);
            streams = Converter.ConvertToImages(MyDir + "Big document.docx", imageSaveOptions);

            streams = Converter.ConvertToImages(new Document(MyDir + "Big document.docx"), SaveFormat.Png);

            streams = Converter.ConvertToImages(new Document(MyDir + "Big document.docx"), imageSaveOptions);
            //ExEnd:ConvertToImagesStream
        }

        [Test]
        public void ConvertToImagesFromStream()
        {
            //ExStart:ConvertToImagesFromStream
            //GistId:708ce40a68fac5003d46f6b4acfd5ff1
            //ExFor:Converter.ConvertToImages(Stream, SaveFormat)
            //ExFor:Converter.ConvertToImages(Stream, ImageSaveOptions)
            //ExSummary:Shows how to convert document to images from stream.
            using (FileStream streamIn = new FileStream(MyDir + "Big document.docx", FileMode.Open, FileAccess.Read))
            {
                Stream[] streams = Converter.ConvertToImages(streamIn, SaveFormat.Jpeg);

                ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Png);
                imageSaveOptions.PageSet = new PageSet(1);
                streams = Converter.ConvertToImages(streamIn, imageSaveOptions);
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
                    LoadOptions loadOptions = new LoadOptions() { Password = "{Asp0se}P@ssw0rd" };
                    SaveTo(docName, loadOptions, new PdfSaveOptions(), "pdf");
                    AssertResult("pdf");

                    break;

                case "HTML":
                    HtmlFixedSaveOptions htmlSaveOptions = new HtmlFixedSaveOptions() { PageSet = new PageSet(0) };
                    SaveTo(docName, new LoadOptions(), htmlSaveOptions, "html");
                    AssertResult("html");

                    break;

                case "XPS":
                    SaveTo(docName, new LoadOptions(), new XpsSaveOptions(), "xps");
                    AssertResult("xps");

                    break;

                case "JPEG":
                    ImageSaveOptions jpegSaveOptions = new ImageSaveOptions(SaveFormat.Jpeg) { JpegQuality = 10 };
                    SaveTo(docName, new LoadOptions(), jpegSaveOptions, "jpeg");
                    AssertResult("jpeg");

                    break;

                case "PNG":
                    ImageSaveOptions pngSaveOptions = new ImageSaveOptions(SaveFormat.Png)
                    {
                        PageSet = new PageSet(0, 1),
                        JpegQuality = 50
                    };
                    SaveTo(docName, new LoadOptions(), pngSaveOptions, "png");
                    AssertResult("png");

                    break;

                case "TIFF":
                    ImageSaveOptions tiffSaveOptions = new ImageSaveOptions(SaveFormat.Tiff) { JpegQuality = 100 };
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
                Stream stream = new MemoryStream();
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

                if (imagesStream.Count != 0)
                {
                    for (int i = 0; i < imagesStream.Count; i++)
                    {
                        using (FileStream resultDoc = new FileStream(ArtifactsDir + $"PdfRenderer_{i}.{fileExt}", FileMode.Create))
                            imagesStream[i].CopyTo(resultDoc);
                    }
                }
                else
                {
                    using (FileStream resultDoc = new FileStream(ArtifactsDir + $"PdfRenderer.{fileExt}", FileMode.Create))
                        stream.CopyTo(resultDoc);
                }
            }
        }

        private void AssertResult(string fileExt)
        {
            if (fileExt == "jpeg" || fileExt == "png" || fileExt == "tiff" || fileExt == "bmp")
            {
                Regex reg = new Regex("PdfRenderer_*");

                var images = Directory.GetFiles(ArtifactsDir, $"*.{fileExt}")
                                     .Where(path => reg.IsMatch(path))
                                     .ToList();

                if (fileExt == "png")
                    Assert.AreEqual(2, images.Count);
                else
                    Assert.AreEqual(5, images.Count);
            }
            else
            {
                if (fileExt == "xps")
                {
                    var doc = new XpsDocument(ArtifactsDir + $"PdfRenderer.{fileExt}");
                    AssertXpsText(doc);
                }
                else
                {
                    var doc = new Document(ArtifactsDir + $"PdfRenderer.{fileExt}");
                    var content = doc.GetText().Replace("\r", " ");

                    Assert.True(content.Contains("Heading 1 Heading 1.1.1.1 Heading 1.1.1.2"));
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
                Assert.True(new[] { "Heading 1", "Head", "ing 1" }.Any(c => ((XpsGlyphs)element).UnicodeString.Contains(c)));
        }

        [Test]
        public void CompareDocuments()
        {
            //ExStart:CompareDocuments
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Comparer.Compare(String, String, String, String, DateTime)
            //ExFor:Comparer.Compare(String, String, String, SaveFormat, String, DateTime)
            //ExFor:Comparer.Compare(String, String, String, String, DateTime, CompareOptions)
            //ExFor:Comparer.Compare(String, String, String, SaveFormat, String, DateTime, CompareOptions)
            //ExSummary:Shows how to simple compare documents.
            // There is a several ways to compare documents:
            string firstDoc = MyDir + "Table column bookmarks.docx";
            string secondDoc = MyDir + "Table column bookmarks.doc";

            Comparer.Compare(firstDoc, secondDoc, ArtifactsDir + "LowCode.CompareDocuments.1.docx", "Author", new DateTime());
            Comparer.Compare(firstDoc, secondDoc, ArtifactsDir + "LowCode.CompareDocuments.2.docx", SaveFormat.Docx, "Author", new DateTime());
            Comparer.Compare(firstDoc, secondDoc, ArtifactsDir + "LowCode.CompareDocuments.3.docx", "Author", new DateTime(), new CompareOptions() { IgnoreCaseChanges = true });
            Comparer.Compare(firstDoc, secondDoc, ArtifactsDir + "LowCode.CompareDocuments.4.docx", SaveFormat.Docx, "Author", new DateTime(), new CompareOptions() { IgnoreCaseChanges = true });
            //ExEnd:CompareDocuments
        }

        [Test]
        public void CompareStreamDocuments()
        {
            //ExStart:CompareStreamDocuments
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Comparer.Compare(Stream, Stream, Stream, SaveFormat, String, DateTime)
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
                        Comparer.Compare(firstStreamIn, secondStreamIn, streamOut, SaveFormat.Docx, "Author", new DateTime(), new CompareOptions() { IgnoreCaseChanges = true });
                }
            }
            //ExEnd:CompareStreamDocuments
        }

        [Test]
        public void MailMerge()
        {
            //ExStart:MailMerge
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:MailMerger.Execute(String, String, String[], Object[])
            //ExFor:MailMerger.Execute(String, String, SaveFormat, String[], Object[])
            //ExFor:MailMerger.Execute(String, String, SaveFormat, MailMergeOptions, String[], Object[])
            //ExSummary:Shows how to do mail merge operation for a single record.
            // There is a several ways to do mail merge operation:
            string doc = MyDir + "Mail merge.doc";

            string[] fieldNames = new string[] { "FirstName", "Location", "SpecialCharsInName()" };
            string[] fieldValues = new string[] { "James Bond", "London", "Classified" };

            MailMerger.Execute(doc, ArtifactsDir + "LowCode.MailMerge.1.docx", fieldNames, fieldValues);
            MailMerger.Execute(doc, ArtifactsDir + "LowCode.MailMerge.2.docx", SaveFormat.Docx, fieldNames, fieldValues);
            MailMerger.Execute(doc, ArtifactsDir + "LowCode.MailMerge.3.docx", SaveFormat.Docx, new MailMergeOptions() { TrimWhitespaces = true }, fieldNames, fieldValues);
            //ExEnd:MailMerge
        }

        [Test]
        public void MailMergeStream()
        {
            //ExStart:MailMergeStream
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, String[], Object[])
            //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, MailMergeOptions, String[], Object[])
            //ExSummary:Shows how to do mail merge operation for a single record from the stream.
            // There is a several ways to do mail merge operation using documents from the stream:
            string[] fieldNames = new string[] { "FirstName", "Location", "SpecialCharsInName()" };
            string[] fieldValues = new string[] { "James Bond", "London", "Classified" };

            using (FileStream streamIn = new FileStream(MyDir + "Mail merge.doc", FileMode.Open, FileAccess.Read))
            {
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MailMergeStream.1.docx", FileMode.Create, FileAccess.ReadWrite))
                    MailMerger.Execute(streamIn, streamOut, SaveFormat.Docx, fieldNames, fieldValues);

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MailMergeStream.2.docx", FileMode.Create, FileAccess.ReadWrite))
                    MailMerger.Execute(streamIn, streamOut, SaveFormat.Docx, new MailMergeOptions() { TrimWhitespaces = true }, fieldNames, fieldValues);
            }
            //ExEnd:MailMergeStream
        }

        [Test]
        public void MailMergeDataRow()
        {
            //ExStart:MailMergeDataRow
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:MailMerger.Execute(String, String, DataRow)
            //ExFor:MailMerger.Execute(String, String, SaveFormat, DataRow)
            //ExFor:MailMerger.Execute(String, String, SaveFormat, MailMergeOptions, DataRow)
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
            MailMerger.Execute(doc, ArtifactsDir + "LowCode.MailMergeDataRow.3.docx", SaveFormat.Docx, new MailMergeOptions() { TrimWhitespaces = true }, dataRow);
            //ExEnd:MailMergeDataRow
        }

        [Test]
        public void MailMergeStreamDataRow()
        {
            //ExStart:MailMergeStreamDataRow
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, DataRow)
            //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, MailMergeOptions, DataRow)
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
                    MailMerger.Execute(streamIn, streamOut, SaveFormat.Docx, new MailMergeOptions() { TrimWhitespaces = true }, dataRow);
            }
            //ExEnd:MailMergeStreamDataRow
        }

        [Test]
        public void MailMergeDataTable()
        {
            //ExStart:MailMergeDataTable
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:MailMerger.Execute(String, String, DataTable)
            //ExFor:MailMerger.Execute(String, String, SaveFormat, DataTable)
            //ExFor:MailMerger.Execute(String, String, SaveFormat, MailMergeOptions, DataTable)
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
            MailMerger.Execute(doc, ArtifactsDir + "LowCode.MailMergeDataTable.3.docx", SaveFormat.Docx, new MailMergeOptions() { TrimWhitespaces = true }, dataTable);
            //ExEnd:MailMergeDataTable
        }

        [Test]
        public void MailMergeStreamDataTable()
        {
            //ExStart:MailMergeStreamDataTable
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, DataTable)
            //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, MailMergeOptions, DataTable)
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
                    MailMerger.Execute(streamIn, streamOut, SaveFormat.Docx, new MailMergeOptions() { TrimWhitespaces = true }, dataTable);
            }
            //ExEnd:MailMergeStreamDataTable
        }

        [Test]
        public void MailMergeWithRegionsDataTable()
        {
            //ExStart:MailMergeWithRegionsDataTable
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:MailMerger.ExecuteWithRegions(String, String, DataTable)
            //ExFor:MailMerger.ExecuteWithRegions(String, String, SaveFormat, DataTable)
            //ExFor:MailMerger.ExecuteWithRegions(String, String, SaveFormat, MailMergeOptions, DataTable)
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
            MailMerger.ExecuteWithRegions(doc, ArtifactsDir + "LowCode.MailMergeWithRegionsDataTable.3.docx", SaveFormat.Docx, new MailMergeOptions() { TrimWhitespaces = true }, dataTable);
            //ExEnd:MailMergeWithRegionsDataTable
        }

        [Test]
        public void MailMergeStreamWithRegionsDataTable()
        {
            //ExStart:MailMergeStreamWithRegionsDataTable
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:MailMerger.ExecuteWithRegions(Stream, Stream, SaveFormat, DataTable)
            //ExFor:MailMerger.ExecuteWithRegions(Stream, Stream, SaveFormat, MailMergeOptions, DataTable)
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
                    MailMerger.ExecuteWithRegions(streamIn, streamOut, SaveFormat.Docx, new MailMergeOptions() { TrimWhitespaces = true }, dataTable);
            }
            //ExEnd:MailMergeStreamWithRegionsDataTable
        }

        [Test]
        public void MailMergeWithRegionsDataSet()
        {
            //ExStart:MailMergeWithRegionsDataSet
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:MailMerger.ExecuteWithRegions(String, String, DataSet)
            //ExFor:MailMerger.ExecuteWithRegions(String, String, SaveFormat, DataSet)
            //ExFor:MailMerger.ExecuteWithRegions(String, String, SaveFormat, MailMergeOptions, DataSet)
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
            MailMerger.ExecuteWithRegions(doc, ArtifactsDir + "LowCode.MailMergeWithRegionsDataSet.3.docx", SaveFormat.Docx, new MailMergeOptions() { TrimWhitespaces = true }, dataSet);
            //ExEnd:MailMergeWithRegionsDataSet
        }

        [Test]
        public void MailMergeStreamWithRegionsDataSet()
        {
            //ExStart:MailMergeStreamWithRegionsDataSet
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:MailMerger.ExecuteWithRegions(Stream, Stream, SaveFormat, DataSet)
            //ExFor:MailMerger.ExecuteWithRegions(Stream, Stream, SaveFormat, MailMergeOptions, DataSet)
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
                    MailMerger.ExecuteWithRegions(streamIn, streamOut, SaveFormat.Docx, new MailMergeOptions() { TrimWhitespaces = true }, dataSet);
            }
            //ExEnd:MailMergeStreamWithRegionsDataSet
        }

        [Test]
        public void Replace()
        {
            //ExStart:Replace
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Replacer.Replace(String, String, String, String)
            //ExFor:Replacer.Replace(String, String, SaveFormat, String, String)
            //ExFor:Replacer.Replace(String, String, SaveFormat, String, String, FindReplaceOptions)
            //ExSummary:Shows how to replace string in the document.
            // There is a several ways to replace string in the document:
            string doc = MyDir + "Footer.docx";
            string pattern = "(C)2006 Aspose Pty Ltd.";
            string replacement = "Copyright (C) 2024 by Aspose Pty Ltd.";

            Replacer.Replace(doc, ArtifactsDir + "LowCode.Replace.1.docx", pattern, replacement);
            Replacer.Replace(doc, ArtifactsDir + "LowCode.Replace.2.docx", SaveFormat.Docx, pattern, replacement);
            Replacer.Replace(doc, ArtifactsDir + "LowCode.Replace.3.docx", SaveFormat.Docx, pattern, replacement, new FindReplaceOptions() { FindWholeWordsOnly = false });
            //ExEnd:Replace
        }

        [Test]
        public void ReplaceStream()
        {
            //ExStart:ReplaceStream
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Replacer.Replace(Stream, Stream, SaveFormat, String, String)
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
                    Replacer.Replace(streamIn, streamOut, SaveFormat.Docx, pattern, replacement, new FindReplaceOptions() { FindWholeWordsOnly = false });
            }
            //ExEnd:ReplaceStream
        }

        [Test]
        public void ReplaceRegex()
        {
            //ExStart:ReplaceRegex
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Replacer.Replace(String, String, Regex, String)
            //ExFor:Replacer.Replace(String, String, SaveFormat, Regex, String)
            //ExFor:Replacer.Replace(String, String, SaveFormat, Regex, String, FindReplaceOptions)
            //ExSummary:Shows how to replace string with regex in the document.
            // There is a several ways to replace string with regex in the document:
            string doc = MyDir + "Footer.docx";
            Regex pattern = new Regex("gr(a|e)y");
            string replacement = "lavender";

            Replacer.Replace(doc, ArtifactsDir + "LowCode.ReplaceRegex.1.docx", pattern, replacement);
            Replacer.Replace(doc, ArtifactsDir + "LowCode.ReplaceRegex.2.docx", SaveFormat.Docx, pattern, replacement);
            Replacer.Replace(doc, ArtifactsDir + "LowCode.ReplaceRegex.3.docx", SaveFormat.Docx, pattern, replacement, new FindReplaceOptions() { FindWholeWordsOnly = false });
            //ExEnd:ReplaceRegex
        }

        [Test]
        public void ReplaceStreamRegex()
        {
            //ExStart:ReplaceStreamRegex
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Replacer.Replace(Stream, Stream, SaveFormat, Regex, String)
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
                    Replacer.Replace(streamIn, streamOut, SaveFormat.Docx, pattern, replacement, new FindReplaceOptions() { FindWholeWordsOnly = false });
            }
            //ExEnd:ReplaceStreamRegex
        }

        //ExStart:BuildReportData
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:ReportBuilder.BuildReport(String, String, Object)
        //ExFor:ReportBuilder.BuildReport(String, String, Object, ReportBuilderOptions)
        //ExFor:ReportBuilder.BuildReport(String, String, SaveFormat, Object)
        //ExFor:ReportBuilder.BuildReport(String, String, SaveFormat, Object, ReportBuilderOptions)
        //ExSummary:Shows how to populate document with data.
        [Test] //ExSkip
        public void BuildReportData()
        {
            // There is a several ways to populate document with data:
            string doc = MyDir + "Reporting engine template - If greedy.docx";

            AsposeData obj = new AsposeData { List = new List<string> { "abc" } };

            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportWithObject.1.docx", obj);
            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportWithObject.2.docx", obj, new ReportBuilderOptions() { Options = ReportBuildOptions.AllowMissingMembers });
            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportWithObject.3.docx", SaveFormat.Docx, obj);
            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportWithObject.4.docx", SaveFormat.Docx, obj, new ReportBuilderOptions() { Options = ReportBuildOptions.AllowMissingMembers });
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
            //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object)
            //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object, ReportBuilderOptions)
            //ExSummary:Shows how to populate document with data using documents from the stream.
            // There is a several ways to populate document with data using documents from the stream:
            AsposeData obj = new AsposeData { List = new List<string> { "abc" } };

            using (FileStream streamIn = new FileStream(MyDir + "Reporting engine template - If greedy.docx", FileMode.Open, FileAccess.Read))
            {
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.BuildReportDataStream.1.docx", FileMode.Create, FileAccess.ReadWrite))
                    ReportBuilder.BuildReport(streamIn, streamOut, SaveFormat.Docx, obj);

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.BuildReportDataStream.2.docx", FileMode.Create, FileAccess.ReadWrite))
                    ReportBuilder.BuildReport(streamIn, streamOut, SaveFormat.Docx, obj, new ReportBuilderOptions() { Options = ReportBuildOptions.AllowMissingMembers });
            }
            //ExEnd:BuildReportDataStream
        }

        //ExStart:BuildReportDataSource
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:ReportBuilder.BuildReport(String, String, Object, String)
        //ExFor:ReportBuilder.BuildReport(String, String, Object[], String[])
        //ExFor:ReportBuilder.BuildReport(String, String, Object, String, ReportBuilderOptions)
        //ExFor:ReportBuilder.BuildReport(String, String, SaveFormat, Object, String)
        //ExFor:ReportBuilder.BuildReport(String, String, SaveFormat, Object, String, ReportBuilderOptions)
        //ExSummary:Shows how to populate document with data sources.
        [Test] //ExSkip
        public void BuildReportDataSource()
        {
            // There is a several ways to populate document with data sources:
            string doc = MyDir + "Report building.docx";

            MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "Hello World");

            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportDataSource.1.docx", sender, "s");
            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportDataSource.2.docx", new object[] { sender }, new[] { "s" });
            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportDataSource.3.docx", sender, "s", new ReportBuilderOptions() { Options = ReportBuildOptions.AllowMissingMembers });
            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportDataSource.4.docx", SaveFormat.Docx, sender, "s");
            ReportBuilder.BuildReport(doc, ArtifactsDir + "LowCode.BuildReportDataSource.5.docx", SaveFormat.Docx, sender, "s", new ReportBuilderOptions() { Options = ReportBuildOptions.AllowMissingMembers });
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
            //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object[], String[])
            //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object, String)
            //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object, String, ReportBuilderOptions)
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
                    ReportBuilder.BuildReport(streamIn, streamOut, SaveFormat.Docx, sender, "s", new ReportBuilderOptions() { Options = ReportBuildOptions.AllowMissingMembers });
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
            //ExFor:Splitter.Split(String, String, SplitOptions)
            //ExFor:Splitter.Split(String, String, SaveFormat, SplitOptions)
            //ExSummary:Shows how to split document by pages.
            string doc = MyDir + "Big document.docx";

            Splitter.Split(doc, ArtifactsDir + "LowCode.SplitDocument.1.docx", new SplitOptions() { SplitCriteria = SplitCriteria.Page });
            Splitter.Split(doc, ArtifactsDir + "LowCode.SplitDocument.2.docx", SaveFormat.Docx, new SplitOptions() { SplitCriteria = SplitCriteria.Page });
            //ExEnd:SplitDocument
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
                Stream[] stream = Splitter.Split(streamIn, SaveFormat.Docx, new SplitOptions() { SplitCriteria = SplitCriteria.Page });
            }
            //ExEnd:SplitDocumentStream
        }

        [Test]
        public void WatermarkText()
        {
            //ExStart:WatermarkText
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Watermarker.SetText(String, String, String)
            //ExFor:Watermarker.SetText(String, String, SaveFormat, String)
            //ExFor:Watermarker.SetText(String, String, String, TextWatermarkOptions)
            //ExFor:Watermarker.SetText(String, String, SaveFormat, String, TextWatermarkOptions)
            //ExSummary:Shows how to insert watermark text to the document.
            string doc = MyDir + "Big document.docx";
            string watermarkText = "This is a watermark";

            Watermarker.SetText(doc, ArtifactsDir + "LowCode.WatermarkText.1.docx", watermarkText);
            Watermarker.SetText(doc, ArtifactsDir + "LowCode.WatermarkText.2.docx", SaveFormat.Docx, watermarkText);
            Watermarker.SetText(doc, ArtifactsDir + "LowCode.WatermarkText.3.docx", watermarkText, new TextWatermarkOptions() { Color = Color.Red });
            Watermarker.SetText(doc, ArtifactsDir + "LowCode.WatermarkText.4.docx", SaveFormat.Docx, watermarkText, new TextWatermarkOptions() { Color = Color.Red });
            //ExEnd:WatermarkText
        }

        [Test]
        public void WatermarkTextStream()
        {
            //ExStart:WatermarkTextStream
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Watermarker.SetText(Stream, Stream, SaveFormat, String)
            //ExFor:Watermarker.SetText(Stream, Stream, SaveFormat, String, TextWatermarkOptions)
            //ExSummary:Shows how to insert watermark text to the document from the stream.
            string watermarkText = "This is a watermark";

            using (FileStream streamIn = new FileStream(MyDir + "Document.docx", FileMode.Open, FileAccess.Read))
            {
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.WatermarkTextStream.1.docx", FileMode.Create, FileAccess.ReadWrite))
                    Watermarker.SetText(streamIn, streamOut, SaveFormat.Docx, watermarkText);

                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.WatermarkTextStream.2.docx", FileMode.Create, FileAccess.ReadWrite))
                    Watermarker.SetText(streamIn, streamOut, SaveFormat.Docx, watermarkText, new TextWatermarkOptions() { Color = Color.Red });
            }
            //ExEnd:WatermarkTextStream
        }

        [Test]
        public void WatermarkImage()
        {
            //ExStart:WatermarkImage
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Watermarker.SetImage(String, String, String)
            //ExFor:Watermarker.SetImage(String, String, SaveFormat, String)
            //ExFor:Watermarker.SetImage(String, String, String, ImageWatermarkOptions)
            //ExFor:Watermarker.SetImage(String, String, SaveFormat, String, ImageWatermarkOptions)
            //ExSummary:Shows how to insert watermark image to the document.
            string doc = MyDir + "Document.docx";
            string watermarkImage = ImageDir + "Logo.jpg";

            Watermarker.SetImage(doc, ArtifactsDir + "LowCode.SetWatermarkImage.1.docx", watermarkImage);
            Watermarker.SetImage(doc, ArtifactsDir + "LowCode.SetWatermarkText.2.docx", SaveFormat.Docx, watermarkImage);
            Watermarker.SetImage(doc, ArtifactsDir + "LowCode.SetWatermarkText.3.docx", watermarkImage, new ImageWatermarkOptions() { Scale = 50 });
            Watermarker.SetImage(doc, ArtifactsDir + "LowCode.SetWatermarkText.4.docx", SaveFormat.Docx, watermarkImage, new ImageWatermarkOptions() { Scale = 50 });
            //ExEnd:WatermarkImage
        }

        [Test]
        public void WatermarkImageStream()
        {
            //ExStart:WatermarkImageStream
            //GistId:695136dbbe4f541a8a0a17b3d3468689
            //ExFor:Watermarker.SetImage(Stream, Stream, SaveFormat, Image)
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
                    Watermarker.SetImage(streamIn, streamOut, SaveFormat.Docx, System.Drawing.Image.FromFile(ImageDir + "Logo.jpg"), new ImageWatermarkOptions() { Scale = 50 });
#endif //ExSkip
            }
            //ExEnd:WatermarkImageStream
        }
    }
}
