using Aspose.Words;
using Aspose.Words.Markup;
using NUnit.Framework;

namespace QaTests.Tests
{
    /// <summary>
    /// Tests that verify work with structured document tags in the document 
    /// </summary>
    [TestFixture]
    internal class QaStructuredDocumentTag : QaTestsBase
    {
        [Test]
        public void RepeatingSection()
        {
            Document doc = new Document(MyDir + "TestRepeatingSection.docx");
            NodeCollection sdts = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

            //Assert that the node have sdttype - RepeatingSection and it's not detected as RichText
            StructuredDocumentTag sdt = (StructuredDocumentTag)sdts[0];
            Assert.AreEqual(SdtType.RepeatingSection, sdt.SdtType);

            //Assert that the node have sdttype - RichText 
            sdt = (StructuredDocumentTag)sdts[1];
            Assert.AreNotEqual(SdtType.RepeatingSection, sdt.SdtType);
        }
    }
}
