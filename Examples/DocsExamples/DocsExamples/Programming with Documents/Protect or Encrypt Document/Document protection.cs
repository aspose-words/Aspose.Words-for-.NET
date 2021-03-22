using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Protect_or_Encrypt_Document
{
    class DocumentProtection : DocsExamplesBase
    {
        [Test]
        public void Protect()
        {
            //ExStart:ProtectDocument
            Document doc = new Document(MyDir + "Document.docx");
            doc.Protect(ProtectionType.AllowOnlyFormFields, "password");
            //ExEnd:ProtectDocument
        }

        [Test]
        public void Unprotect()
        {
            //ExStart:UnprotectDocument
            Document doc = new Document(MyDir + "Document.docx");
            doc.Unprotect();
            //ExEnd:UnprotectDocument
        }

        [Test]
        public void GetProtectionType()
        {
            //ExStart:GetProtectionType
            Document doc = new Document(MyDir + "Document.docx");
            ProtectionType protectionType = doc.ProtectionType;
            //ExEnd:GetProtectionType
        }
    }
}