// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class ConvertFromDocmToDocx : TestUtil
    {
        [Test]
        public void ConvertFromDocmToDocxFeature()
        {
            bool fileChanged = false;

            using (WordprocessingDocument document =
                WordprocessingDocument.Open(MyDir + "Convert from docm to docx.docm", true))
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
                if (File.Exists(ArtifactsDir + "Convert from docm to docx - OpenXML.docm"))
                    File.Delete(ArtifactsDir + "Convert from docm to docx - OpenXML.docm");

                File.Move(MyDir + "Convert from docm to docx.docm",
                    ArtifactsDir + "Convert from docm to docx - OpenXML.docm");
            }
        }
    }
}
