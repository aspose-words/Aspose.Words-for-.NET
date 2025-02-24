// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class RemoveSectionBreaks : TestUtil
    {
        [Test]
        public void RemoveSectionBreaksFeature()
        {
            File.Copy(MyDir + "Remove section breaks.docx", ArtifactsDir + "Remove section breaks - OpenXML.docx", true);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(ArtifactsDir + "Remove section breaks - OpenXML.docx", true))
            {
                // Get the main document part.
                var body = doc.MainDocumentPart.Document.Body;

                // Find all section breaks in the document.
                var sectionProperties = body.Descendants<SectionProperties>().ToList();
                // Remove each section properties element.
                foreach (var section in sectionProperties)
                    section.Remove();
            }
        }
    }
}
