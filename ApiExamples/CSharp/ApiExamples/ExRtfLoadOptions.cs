// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExRtfLoadOptions : ApiExampleBase
    {
        [Test]
        public void RecognizeUtf8Text()
        {
            //ExStart
            //ExFor:RtfLoadOptions.RecognizeUtf8Text
            //ExSummary:Shows how to detect UTF8 characters during import.
            RtfLoadOptions loadOptions = new RtfLoadOptions
            {
                RecognizeUtf8Text = true
            };

            Document doc = new Document(MyDir + "RtfLoadOptions.RecognizeUtf8Text.rtf", loadOptions);
            doc.Save(ArtifactsDir + "RtfLoadOptions.RecognizeUtf8Text.rtf");
            //ExEnd
        }
    }
}