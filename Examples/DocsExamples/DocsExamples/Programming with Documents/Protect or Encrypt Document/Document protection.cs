using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Protect_or_Encrypt_Document
{
    class DocumentProtection : DocsExamplesBase
    {
        [Test]
        public void PasswordProtection()
        {
            //ExStart:PasswordProtection
            Document doc = new Document();

            // Apply document protection.
            doc.Protect(ProtectionType.NoProtection, "password");

            doc.Save(ArtifactsDir + "DocumentProtection.PasswordProtection.docx");
            //ExEnd:PasswordProtection
        }

        [Test]
        public void AllowOnlyFormFieldsProtect()
        {
            //ExStart:AllowOnlyFormFieldsProtect
            // Insert two sections with some text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Text added to a document.");

            // A document protection only works when document protection is turned and only editing in form fields is allowed.
            doc.Protect(ProtectionType.AllowOnlyFormFields, "password");

            // Save the protected document.
            doc.Save(ArtifactsDir + "DocumentProtection.AllowOnlyFormFieldsProtect.docx");
            //ExEnd:AllowOnlyFormFieldsProtect
        }

        [Test]
        public void RemoveDocumentProtection()
        {
            //ExStart:RemoveDocumentProtection
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Text added to a document.");

            // Documents can have protection removed either with no password, or with the correct password.
            doc.Unprotect();
            doc.Protect(ProtectionType.ReadOnly, "newPassword");
            doc.Unprotect("newPassword");

            doc.Save(ArtifactsDir + "DocumentProtection.RemoveDocumentProtection.docx");
            //ExEnd:RemoveDocumentProtection
        }

        [Test]
        public void UnrestrictedEditableRegions()
        {
            //ExStart:UnrestrictedEditableRegions
            // Upload a document and make it as read-only.
            Document doc = new Document(MyDir + "Document.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            doc.Protect(ProtectionType.ReadOnly, "MyPassword");

            builder.Writeln("Hello world! Since we have set the document's protection level to read-only, " + "we cannot edit this paragraph without the password.");

            // Start an editable range.
            EditableRangeStart edRangeStart = builder.StartEditableRange();
            // An EditableRange object is created for the EditableRangeStart that we just made.
            EditableRange editableRange = edRangeStart.EditableRange;

            // Put something inside the editable range.
            builder.Writeln("Paragraph inside first editable range");

            // An editable range is well-formed if it has a start and an end.
            EditableRangeEnd edRangeEnd = builder.EndEditableRange();

            builder.Writeln("This paragraph is outside any editable ranges, and cannot be edited.");

            doc.Save(ArtifactsDir + "DocumentProtection.UnrestrictedEditableRegions.docx");
            //ExEnd:UnrestrictedEditableRegions
        }

        [Test]
        public void UnrestrictedSection()
        {
            //ExStart:UnrestrictedSection
            // Insert two sections with some text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Section 1. Unprotected.");
            builder.InsertBreak(BreakType.SectionBreakContinuous);
            builder.Writeln("Section 2. Protected.");

            // Section protection only works when document protection is turned and only editing in form fields is allowed.
            doc.Protect(ProtectionType.AllowOnlyFormFields, "password");

            // By default, all sections are protected, but we can selectively turn protection off.
            doc.Sections[0].ProtectedForForms = false;
            doc.Save(ArtifactsDir + "DocumentProtection.UnrestrictedSection.docx");

            doc = new Document(ArtifactsDir + "DocumentProtection.UnrestrictedSection.docx");
            Assert.False(doc.Sections[0].ProtectedForForms);
            Assert.True(doc.Sections[1].ProtectedForForms);
            //ExEnd:UnrestrictedSection
        }

        [Test]
        public void GetProtectionType()
        {
            //ExStart:GetProtectionType
            Document doc = new Document(MyDir + "Document.docx");
            ProtectionType protectionType = doc.ProtectionType;
            //ExEnd:GetProtectionType
        }

        [Test]
        public void ReadOnlyProtection()
        {
            //ExStart:ReadOnlyProtection
            //GistId:7cf6735e83804ba8942663695b22ee42
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Open document as read-only");

            // Enter a password that's up to 15 characters long.
            doc.WriteProtection.SetPassword("MyPassword");

            // Make the document as read-only.
            doc.WriteProtection.ReadOnlyRecommended = true;

            // Apply write protection as read-only.
            doc.Protect(ProtectionType.ReadOnly);
            doc.Save(ArtifactsDir + "DocumentProtection.ReadOnlyProtection.docx");
            //ExEnd:ReadOnlyProtection
        }

        [Test]
        public void RemoveReadOnlyRestriction()
        {
            //ExStart:RemoveReadOnlyRestriction
            //GistId:7cf6735e83804ba8942663695b22ee42
            Document doc = new Document();
            
            // Enter a password that's up to 15 characters long.
            doc.WriteProtection.SetPassword("MyPassword");

            // Remove the read-only option.
            doc.WriteProtection.ReadOnlyRecommended = false;

            // Apply write protection without any protection.
            doc.Protect(ProtectionType.NoProtection);
            doc.Save(ArtifactsDir + "DocumentProtection.RemoveReadOnlyRestriction.docx");
            //ExEnd:RemoveReadOnlyRestriction
        }
    }
}