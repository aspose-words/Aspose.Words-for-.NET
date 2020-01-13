// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExXpsSaveOptions : ApiExampleBase
    {
        [Test]
        public void OptimizeOutput()
        {
            //ExStart
            //ExFor:FixedPageSaveOptions.OptimizeOutput
            //ExSummary:Shows how to optimize document objects while saving to xps.
            Document doc = new Document(MyDir + "XPSOutputOptimize.docx");

            XpsSaveOptions saveOptions = new XpsSaveOptions { OptimizeOutput = true };

            doc.Save(ArtifactsDir + "XpsSaveOptions.OptimizeOutput.xps", saveOptions);
            //ExEnd
        }
    }
}