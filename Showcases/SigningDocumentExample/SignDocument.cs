using System;

namespace SigningDocumentExample
{
    public class SignDocument
    {
        public Guid DocumentId { get; set; }
        public string DocumentName { get; set; }
        public byte[] Document { get; set; }
    }
}
