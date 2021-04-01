// Copyright (c) Aspose 2002-2021. All Rights Reserved.

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
        public void RemoveHeaderFooterFeature()
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(MyDir + "Document.docx", true))
            {
                var mainDocumentPart = doc.MainDocumentPart;

                // Count the header and footer parts and continue if there are any.
                if (mainDocumentPart.HeaderParts.Any() || mainDocumentPart.FooterParts.Any())
                {
                    // Remove the header and footer parts.
                    mainDocumentPart.DeleteParts(mainDocumentPart.HeaderParts);
                    mainDocumentPart.DeleteParts(mainDocumentPart.FooterParts);

                    // Get a reference to the root element of the main document part.
                    Document document = mainDocumentPart.Document;

                    // Remove all references to the headers and footers.

                    // First, create a list of all descendants of type HeaderReference.
                    // Then, navigate the list and call remove on each item to delete the reference.
                    var headers = document.Descendants<HeaderReference>().ToList();

                    foreach (var header in headers)
                        header.Remove();

                    // First, create a list of all descendants of type FooterReference.
                    // Then, navigate the list and call remove on each item to delete the reference.
                    var footers = document.Descendants<FooterReference>().ToList();

                    foreach (var footer in footers)
                        footer.Remove();

                    using (Stream stream = File.Create(ArtifactsDir + "Remove header and footer - OpenXML.docx"))
                    {
                        document.Save(stream);
                    }
                }
            }
        }
    }
}