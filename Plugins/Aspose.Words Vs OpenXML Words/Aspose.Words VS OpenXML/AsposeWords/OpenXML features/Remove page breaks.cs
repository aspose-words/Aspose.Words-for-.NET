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
    public class RemovePageBreaks : TestUtil
    {
        [Test]
        public void RemovePageBreaksFeature()
        {
            File.Copy(MyDir + "Remove page breaks.docx", ArtifactsDir + "Remove page breaks - OpenXML.docx", true);

            using WordprocessingDocument doc = WordprocessingDocument.Open(ArtifactsDir + "Remove page breaks - OpenXML.docx", true);
            // Get the main document part.
            Body body = doc.MainDocumentPart.Document.Body;

            // Find all page breaks in the document.
            List<Break> breaks = body.Descendants<Break>().ToList();
            foreach (var pageBreak in breaks)
                pageBreak.Remove();
        }
    }
}
