using System.IO;
using Aspose.Words;

namespace ApiExamples.TestData
{
    public class DocumentDataSource
    {
        public DocumentDataSource(Document doc)
        {
            this.Document = doc;
        }

        public DocumentDataSource(Stream stream)
        {
            this.DocumentByStream = stream;
        }

        public DocumentDataSource(byte[] byteDoc)
        {
            this.DocumentByByte = byteDoc;
        }

        public DocumentDataSource(string uriToDoc)
        {
            this.DocumentByUri = uriToDoc;
        }

        public Document Document { get; set; }

        public Stream DocumentByStream { get; set; }

        public byte[] DocumentByByte { get; set; }

        public string DocumentByUri { get; set; }
    }
}