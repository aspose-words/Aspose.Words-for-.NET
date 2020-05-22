// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExRtfLoadOptions : ApiExampleBase
    {
        [Test]
        [TestCase(false)]
        [TestCase(true)]
        public void RecognizeUtf8Text(bool doRecognizeUtb8Text)
        {
            //ExStart
            //ExFor:RtfLoadOptions
            //ExFor:RtfLoadOptions.#ctor
            //ExFor:RtfLoadOptions.RecognizeUtf8Text
            //ExSummary:Shows how to detect UTF8 characters during import.
            RtfLoadOptions loadOptions = new RtfLoadOptions
            {
                RecognizeUtf8Text = doRecognizeUtb8Text
            };

            Document doc = new Document(MyDir + "UTF-8 characters.rtf", loadOptions);

            if (doRecognizeUtb8Text)
                Assert.AreEqual("“John Doe´s list of currency symbols”™\r€, ¢, £, ¥, ¤", doc.FirstSection.Body.GetText().Trim());
            else 
                Assert.AreEqual("â€œJohn DoeÂ´s list of currency symbolsâ€\u009dâ„¢\râ‚¬, Â¢, Â£, Â¥, Â¤", doc.FirstSection.Body.GetText().Trim());
            //ExEnd
        }
    }
}