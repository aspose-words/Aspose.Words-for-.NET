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
    public class RemoveHiddenText : TestUtil
    {
        [Test]
        public void RemoveHiddenTextFeature()
        {
            File.Copy(MyDir + "Remove hidden text.docx", ArtifactsDir + "Remove hidden text - OpenXML.docx", true);

            using WordprocessingDocument doc = WordprocessingDocument.Open(ArtifactsDir + "Remove hidden text - OpenXML.docx", true);
            foreach (var paragraph in doc.MainDocumentPart.Document.Body.Elements<Paragraph>())
            {
                // Iterate through all runs in the paragraph.
                foreach (var run in paragraph.Elements<Run>())
                {
                    // Check if the run has properties
                    var runProperties = run.RunProperties;
                    if (runProperties != null)
                    {
                        // Check if the text is hidden.
                        var hidden = runProperties.Elements<Vanish>().FirstOrDefault();
                        if (hidden != null)
                            // Remove the hidden property to unhide the text.
                            hidden.Remove();
                    }
                }
            }
        }
    }
}
