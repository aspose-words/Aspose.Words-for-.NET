using System.IO;
using Aspose.Words;
using NUnit.Framework;

namespace QaTests.Tests
{
    [TestFixture]
    internal class QaDocument : QaTestsBase
    {
        [Test]
        public void DocumentDefaultStyles()
        {
            Document doc = new Document();

            //Add document-wide defaults parameters
            doc.Styles.DefaultFont.Name = "PMingLiU";
            doc.Styles.DefaultFont.Bold = true;
            
            doc.Styles.DefaultParagraphFormat.SpaceAfter = 20;
            doc.Styles.DefaultParagraphFormat.Alignment = ParagraphAlignment.Right;

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Rtf);

            Assert.IsTrue(doc.Styles.DefaultFont.Bold);
            Assert.AreEqual("PMingLiU", doc.Styles.DefaultFont.Name);
            Assert.AreEqual(20, doc.Styles.DefaultParagraphFormat.SpaceAfter);
            Assert.AreEqual(ParagraphAlignment.Right, doc.Styles.DefaultParagraphFormat.Alignment);
        }
    }
}
