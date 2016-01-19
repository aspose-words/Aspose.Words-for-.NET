using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using NUnit.Framework;

namespace QA_Tests.Tests.Reporting
{
    /// <summary>
    /// Tests that verify ReportingEngine functions
    /// </summary>
    [TestFixture]
    internal class QaEngine : QaTestsBase
    {
        private readonly string _image = TestDir + @"Images\Test_636_852.gif";

        [Test]
        public void StretchImage_fitHeight()
        {
            Document doc = DocumentHelper.CreateTemplateDocumentForReportingEngine("<<image [src.Image] -fitHeight>>");

            ImageStream imageStream = new ImageStream(new FileStream(_image, FileMode.Open, FileAccess.Read));

            BuildReport(doc, imageStream, "src");

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

            ImageStream imageStream = new ImageStream(new FileStream(_image, FileMode.Open, FileAccess.Read));

            BuildReport(doc, imageStream, "src");

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

            ImageStream imageStream = new ImageStream(new FileStream(_image, FileMode.Open, FileAccess.Read));

            BuildReport(doc, imageStream, "src");

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
        [ExpectedException(typeof (InvalidOperationException))]
        public void WithoutAllowMissingDataFields()
        {
            Document doc = new Document();

            DocumentHelper.InsertNewRun(doc, "<<if [value == “true”] >>ok<<else>>Cancel<</if>>");

            DataSet dataSet = new DataSet();
            dataSet.ReadXml(TestDir + "DataSet.xml", XmlReadMode.InferSchema);

            BuildReport(doc, dataSet, "Bad");
        }
        /// <summary>
        /// Assert that the exception from previous test is not repeated with AllowMissingMembers parameter
        /// </summary>
        [Test]
        public void WithAllowMissingDataFields()
        {
            Document doc = new Document();

            DocumentHelper.InsertNewRun(doc, "<<if [value == “true”] >>ok<<else>>Cancel<</if>>");

            DataSet dataSet = new DataSet();
            dataSet.ReadXml(TestDir + "DataSet.xml", XmlReadMode.InferSchema);

            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.AllowMissingMembers;

            engine.BuildReport(doc, dataSet, "Bad");
        }

        private static void BuildReport(Document document, object dataSource, string dataSourceName)
        {
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(document, dataSource, dataSourceName);
        }
    }
}

public class ImageStream
{
    public ImageStream(Stream stream)
    {
        Image = stream;
    }

    public Stream Image { get; }
}


