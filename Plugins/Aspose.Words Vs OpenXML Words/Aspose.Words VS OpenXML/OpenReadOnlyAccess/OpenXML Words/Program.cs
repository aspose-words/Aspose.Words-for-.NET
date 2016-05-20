using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        private static string FilePath = @"..\..\..\..\Sample Files\";
        private static string fileName = FilePath + "OpenReadOnlyAccess.docx";
        static void Main(string[] args)
        {
            OpenWordprocessingDocumentReadonly(fileName);
        }
        public static void OpenWordprocessingDocumentReadonly(string filepath)
        {
            // Open a WordprocessingDocument based on a filepath.
            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Open(filepath, false))
            {
                // Assign a reference to the existing document body.  
                Body body = wordDocument.MainDocumentPart.Document.Body;

                // Attempt to add some text.
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("Append text in body, but text is not saved - OpenWordprocessingDocumentReadonly"));

                // Call Save to generate an exception and show that access is read-only.
                // wordDocument.MainDocumentPart.Document.Save();
            }
        }
    }
}
