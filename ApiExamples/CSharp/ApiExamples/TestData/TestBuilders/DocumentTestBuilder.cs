﻿using System.IO;
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
            mDocument = new Document();
            mDocumentStream = Stream.Null;
            mDocumentBytes = new byte[0];
            mDocumentUri = string.Empty;
        }

        public DocumentTestBuilder WithDocument(Document doc)
        {
            mDocument = doc;
            return this;
        }

        public DocumentTestBuilder WithDocumentStream(Stream stream)
        {
            mDocumentStream = stream;
            return this;
        }

        public DocumentTestBuilder WithDocumentBytes(byte[] docBytes)
        {
            mDocumentBytes = docBytes;
            return this;
        }

        public DocumentTestBuilder WithDocumentUri(string docUri)
        {
            mDocumentUri = docUri;
            return this;
        }

        public DocumentTestClass Build()
        {
            return new DocumentTestClass(mDocument, mDocumentStream, mDocumentBytes, mDocumentUri);
        }
    }
}