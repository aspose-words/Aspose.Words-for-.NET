using System.IO;
using Aspose.Words;

namespace ApiExamples.TestData.TestClasses
{
    public class DocumentTestClass
    {
        public Document Document { get; set; }
        public Stream DocumentStream { get; set; }
        public byte[] DocumentBytes { get; set; }
        public string DocumentUri { get; set; }

        public DocumentTestClass(Document doc, Stream docStream, byte[] docBytes, string docUri)
        {
            this.Document = doc;
            this.DocumentStream = docStream;
            this.DocumentBytes = docBytes;
            this.DocumentUri = docUri;
        }
    }
}
