using System.IO;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    class ExComHelper : ApiExampleBase
    {
        [Test]
        public void ComHelper()
        {
            //ExStart
            //ExFor:ComHelper
            //ExFor:ComHelper.#ctor
            //ExFor:ComHelper.Open(Stream)
            //ExFor:ComHelper.Open(String)
            //ExSummary:Shows how to open documents using the ComHelper class.
            // If you need to open a document within a COM application,
            // you will need to do so using the ComHelper class as instead of the Document constructor
            ComHelper comHelper = new ComHelper();

            // There are two ways of using a ComHelper to open a document
            // 1: Using a filename
            Document doc = comHelper.Open(MyDir + "Document.docx");
            Assert.AreEqual("Hello World!\f", doc.GetText());

            // 2: Using a Stream
            using (FileStream stream = new FileStream(MyDir + "Document.docx", FileMode.Open))
            {
                doc = comHelper.Open(stream);
                Assert.AreEqual("Hello World!\f", doc.GetText());
            }
            //ExEnd
        }
    }
}
