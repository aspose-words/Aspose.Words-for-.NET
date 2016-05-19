using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string File = FilePath + "ConvertFromDOCMtoDOCX - OpenXML.docm";
            string NewFile = FilePath + "ConvertFromDOCMtoDOCX - OpenXML - Output.docx";
            ConvertDOCMtoDOCX(File, NewFile);
        }
        // Given a .docm file (with macro storage), remove the VBA 
        // project, reset the document type, and save the document with a new name.
        public static void ConvertDOCMtoDOCX(string oldfileName, string newfileName)
        {
            bool fileChanged = false;

            using (WordprocessingDocument document =
                WordprocessingDocument.Open(oldfileName, true))
            {
                // Access the main document part.
                var docPart = document.MainDocumentPart;

                // Look for the vbaProject part. If it is there, delete it.
                var vbaPart = docPart.VbaProjectPart;
                if (vbaPart != null)
                {
                    // Delete the vbaProject part and then save the document.
                    docPart.DeletePart(vbaPart);
                    docPart.Document.Save();

                    // Change the document type to
                    // not macro-enabled.
                    document.ChangeDocumentType(
                        WordprocessingDocumentType.Document);

                    // Track that the document has been changed.
                    fileChanged = true;
                }
            }

            // If anything goes wrong in this file handling,
            // the code will raise an exception back to the caller.
            if (fileChanged)
            {
                // Create the new .docx filename.

                // If it already exists, it will be deleted!
                if (File.Exists(newfileName))
                {
                    File.Delete(newfileName);
                }

                // Rename the file.
                File.Move(oldfileName, newfileName);
            }
        }
    }
}
