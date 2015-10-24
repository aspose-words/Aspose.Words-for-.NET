// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Words;

namespace Aspose_Word
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = "Change or Replace Header and footer.docx";
            ChangeHeader(path);

        }
        public static void ChangeHeader(string documentPath)
        {
            Document doc = new Document(documentPath);
            DocumentBuilder builder = new DocumentBuilder(doc);

            // --- Create header ---
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Specify header title for the first page.
            builder.Write("Aspose.Words Header");

            // --- Create footer for pages other than first. ---
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

            // Specify Footer text.
            builder.Write("Aspose.Words Footer");

            // Save the resulting document.
            doc.Save(documentPath);
        }
    }
}
