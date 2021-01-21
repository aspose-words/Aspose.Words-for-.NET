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

        // Given a .docm file (with macro storage): remove the VBA 
        // project, reset the document type, and then save the document in the local file system under a new filename.
        public static void ConvertDOCMtoDOCX(string oldfileName, string newfileName)
        {
            bool fileChanged = false;

            using (WordprocessingDocument document =
                WordprocessingDocument.Open(oldfileName, true))
            {
                var docPart = document.MainDocumentPart;

                // Look for the vbaProject part. If it is there, delete it.
                var vbaPart = docPart.VbaProjectPart;
                if (vbaPart != null)
                {
                    // Delete the vbaProject part and then save the document.
                    docPart.DeletePart(vbaPart);
                    docPart.Document.Save();

                    // Change the document type to not macro-enabled.
                    document.ChangeDocumentType(
                        WordprocessingDocumentType.Document);

                    fileChanged = true;
                }
            }

            // If anything goes wrong in this file handling,
            // the code will raise an exception back to the caller.
            if (fileChanged)
            {
                if (File.Exists(newfileName))
                    File.Delete(newfileName);

                File.Move(oldfileName, newfileName);
            }
        }
    }
}
