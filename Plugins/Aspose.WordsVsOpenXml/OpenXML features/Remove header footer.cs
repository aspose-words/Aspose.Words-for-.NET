// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class RemoveHeaderFooter : TestUtil
    {
        [Test]
        public void RemoveHeaderFooterOpenXml()
        {
            //ExStart:RemoveHeaderFooterOpenXml
            //GistId:3fbf1435ff3b2e08f9968067e177307d
            File.Copy(MyDir + "Document.docx", ArtifactsDir + "Remove header and footer - OpenXML.docx", true);

            using WordprocessingDocument doc = WordprocessingDocument.Open(ArtifactsDir + "Remove header and footer - OpenXML.docx", true);

            MainDocumentPart mainDocumentPart = doc.MainDocumentPart;
            // Count the header and footer parts and continue if there are any.
            if (mainDocumentPart.HeaderParts.Any() || mainDocumentPart.FooterParts.Any())
            {
                // Remove the header and footer parts.
                mainDocumentPart.DeleteParts(mainDocumentPart.HeaderParts);
                mainDocumentPart.DeleteParts(mainDocumentPart.FooterParts);

                // Get a reference to the root element of the main document part.
                Document document = mainDocumentPart.Document;

                // First, create a list of all descendants of type HeaderReference.
                // Then, navigate the list and call remove on each item to delete the reference.
                List<HeaderReference> headers = document.Descendants<HeaderReference>().ToList();
                foreach (HeaderReference header in headers)
                    header.Remove();

                // First, create a list of all descendants of type FooterReference.
                // Then, navigate the list and call remove on each item to delete the reference.
                List<FooterReference> footers = document.Descendants<FooterReference>().ToList();
                foreach (FooterReference footer in footers)
                    footer.Remove();
                //ExEnd:RemoveHeaderFooterOpenXml
            }
        }
    }
}