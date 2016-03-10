using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using NUnit.Framework;

namespace QaTests.Tests
{
    /// <summary>
    /// Tests that verify ReportingEngine functions
    /// </summary>
    [TestFixture]
    internal class QaEngine : QaTestsBase
    {
        private readonly string _image = MyDir + @"Images\Test_636_852.gif";

        [Test]
        public void StretchImage_fitHeight()
        {
            Document doc = DocumentHelper.CreateTemplateDocumentForReportingEngine("<<image [src.Image] -fitHeight>>");

            ImageStream imageStream = new ImageStream(new FileStream(this._image, FileMode.Open, FileAccess.Read));

            BuildReport(doc, imageStream, "src", ReportBuildOptions.None);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            doc = new Document(dstStream);

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes)
            {
                // Assert that the image is really insert in textbox 
                Assert.IsTrue(shape.ImageData.HasImage);
                
                //Assert that width is keeped and height is changed
                Assert.AreNotEqual(346.35, shape.Height);
                Assert.AreEqual(431.5, shape.Width);
            }

            dstStream.Dispose();
        }

        [Test]
        public void StretchImage_fitWidth()
        {
            Document doc = DocumentHelper.CreateTemplateDocumentForReportingEngine("<<image [src.Image] -fitWidth>>");

            ImageStream imageStream = new ImageStream(new FileStream(this._image, FileMode.Open, FileAccess.Read));

            BuildReport(doc, imageStream, "src", ReportBuildOptions.None);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            doc = new Document(dstStream);

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes)
            {
                // Assert that the image is really insert in textbox and 
                Assert.IsTrue(shape.ImageData.HasImage);

                //Assert that height is keeped and width is changed
                Assert.AreNotEqual(431.5, shape.Width);
                Assert.AreEqual(346.35, shape.Height);
            }

            dstStream.Dispose();
        }

        [Test]
        public void StretchImage_fitSize()
        {
            Document doc = DocumentHelper.CreateTemplateDocumentForReportingEngine("<<image [src.Image] -fitSize>>");

            ImageStream imageStream = new ImageStream(new FileStream(this._image, FileMode.Open, FileAccess.Read));

            BuildReport(doc, imageStream, "src", ReportBuildOptions.None);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            doc = new Document(dstStream);

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes)
            {
                // Assert that the image is really insert in textbox 
                Assert.IsTrue(shape.ImageData.HasImage);

                //Assert that height is changed and width is changed
                Assert.AreNotEqual(346.35, shape.Height);
                Assert.AreNotEqual(431.5, shape.Width);
            }

            dstStream.Dispose();
        }

        [Test]
        [ExpectedException(typeof(InvalidOperationException))]
        public void WithoutMissingMembers()
        {
            DocumentBuilder builder = new DocumentBuilder();

            //Add templete to the document for reporting engine
            DocumentHelper.InsertBuilderText(builder, new[] { "<<[missingObject.First().id]>>", "<<foreach [in missingObject]>><<[id]>><</foreach>>" });

            //Assert that build report failed without "ReportBuildOptions.AllowMissingMembers"
            BuildReport(builder.Document, new DataSet(), "", ReportBuildOptions.None);
        }

        [Test]
        public void WithMissingMembers()
        {
            DocumentBuilder builder = new DocumentBuilder();

            //Add templete to the document for reporting engine
            DocumentHelper.InsertBuilderText(builder, new[] { "<<[missingObject.First().id]>>", "<<foreach [in missingObject]>><<[id]>><</foreach>>" });

            BuildReport(builder.Document, new DataSet(), "", ReportBuildOptions.AllowMissingMembers);

            //Assert that build report success with "ReportBuildOptions.AllowMissingMembers"
            Assert.AreEqual(
            ControlChar.ParagraphBreak + ControlChar.ParagraphBreak + ControlChar.SectionBreak,
            builder.Document.GetText());
        }

        private static void BuildReport(Document document, object dataSource, string dataSourceName, ReportBuildOptions reportBuildOptions)
        {
            ReportingEngine engine = new ReportingEngine();
            engine.Options = reportBuildOptions;

            engine.BuildReport(document, dataSource, dataSourceName);
        }
    }
}

public class ImageStream
{
    public ImageStream(Stream stream)
    {
        this.Image = stream;
    }

    public Stream Image { get; set; }
}


