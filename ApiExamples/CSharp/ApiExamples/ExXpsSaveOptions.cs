// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExXpsSaveOptions : ApiExampleBase
    {
        [TestCase(false)]
        [TestCase(true)]
        public void OptimizeOutput(bool optimizeOutput)
        {
            //ExStart
            //ExFor:FixedPageSaveOptions.OptimizeOutput
            //ExSummary:Shows how to optimize document objects while saving to xps.
            Document doc = new Document(MyDir + "Unoptimized document.docx");

            // When saving to .xps, we can use SaveOptions to optimize the output in some cases
            XpsSaveOptions saveOptions = new XpsSaveOptions { OptimizeOutput = optimizeOutput };

            doc.Save(ArtifactsDir + "XpsSaveOptions.OptimizeOutput.xps", saveOptions);

            // The input document had adjacent runs with the same formatting, which, if output optimization was enabled,
            // have been combined to save space
            FileInfo outFileInfo = new FileInfo(ArtifactsDir + "XpsSaveOptions.OptimizeOutput.xps");

            if (optimizeOutput)
                Assert.True(outFileInfo.Length < 45000);
            else
                Assert.True(outFileInfo.Length > 60000);
            //ExEnd

            TestUtil.DocPackageFileContainsString(
                optimizeOutput
                    ? "Glyphs OriginX=\"34.294998169\" OriginY=\"10.31799984\" " +
                      "UnicodeString=\"This document contains complex content which can be optimized to save space when \""
                    : "<Glyphs OriginX=\"34.294998169\" OriginY=\"10.31799984\" UnicodeString=\"This\"",
                ArtifactsDir + "XpsSaveOptions.OptimizeOutput.xps", "1.fpage");
        }
    }
}