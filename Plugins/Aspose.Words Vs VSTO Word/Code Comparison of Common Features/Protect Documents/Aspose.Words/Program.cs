namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");

            // Save the document to the local file system with no protection.
            // Anybody who opens this document will freely be able to edit it.
            doc.Save("Protect Documents.Unprotected.docx");

            // Set a read-only type protection with a password.
            doc.Protect(ProtectionType.ReadOnly, "myPassword");

            // Microsoft Word will restrict the editing of this document when we open it.
            // In order to edit it, we will need to press "Stop Protection" in the "Restrict Editing" menu,
            // and then type in the password that we used to protect this document above.
            doc.Save("Protect Documents.Protected.docx");
        }
    }
}
