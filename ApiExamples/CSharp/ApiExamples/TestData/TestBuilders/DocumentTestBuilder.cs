using System.IO;
using ApiExamples.TestData.TestClasses;
using Aspose.Words;

namespace ApiExamples.TestData.TestBuilders
{
    public class DocumentTestBuilder : ApiExampleBase
    {
        private Document mDocument;
        private Stream mDocumentStream;
        private byte[] mDocumentBytes;
        private string mDocumentUri;

        public DocumentTestBuilder()
        {
            this.mDocument = new Document();
            this.mDocumentStream = Stream.Null;
            this.mDocumentBytes = new byte[0];
            this.mDocumentUri = string.Empty;
        }

        public DocumentTestBuilder WithDocument(Document doc)
        {
            this.mDocument = doc;
            return this;
        }

        public DocumentTestBuilder WithDocumentStream(Stream stream)
        {
            this.mDocumentStream = stream;
            return this;
        }

        public DocumentTestBuilder WithDocumentBytes(byte[] docBytes)
        {
            this.mDocumentBytes = docBytes;
            return this;
        }

        public DocumentTestBuilder WithDocumentUri(string docUri)
        {
            this.mDocumentUri = docUri;
            return this;
        }

        public DocumentTestClass Build()
        {
            return new DocumentTestClass(mDocument, mDocumentStream, mDocumentBytes, mDocumentUri);
        }
    }
}
