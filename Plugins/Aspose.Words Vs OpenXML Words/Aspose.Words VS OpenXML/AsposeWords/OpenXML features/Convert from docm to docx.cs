// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

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
        public void DocmToDocxConversion()
        {
            string docmFilePath = MyDir + "Docm to Docx conversion.docm";
            string docxFilePath = ArtifactsDir + "Docm to Docx conversion - OpenXML.docx";

            using (WordprocessingDocument docm = WordprocessingDocument.Open(docmFilePath, false))
            {
                // Create a copy of the .docm file as .docx.
                File.Copy(docmFilePath, docxFilePath, true);

                // Open the new .docx file and remove the macros.
                using (WordprocessingDocument docx = WordprocessingDocument.Open(docxFilePath, true))
                {
                    // Remove the VBA project part (macros).
                    var vbaPart = docx.MainDocumentPart.VbaProjectPart;
                    if (vbaPart != null)
                        docx.MainDocumentPart.DeletePart(vbaPart);

                    // Change the document type to .docx (no macros).
                    docx.ChangeDocumentType(WordprocessingDocumentType.Document);
                }
            }
        }
    }
}
