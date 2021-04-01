// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Loading;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExRtfLoadOptions : ApiExampleBase
    {
        [TestCase(false)]
        [TestCase(true)]
        public void RecognizeUtf8Text(bool recognizeUtf8Text)
        {
            //ExStart
            //ExFor:RtfLoadOptions
            //ExFor:RtfLoadOptions.#ctor
            //ExFor:RtfLoadOptions.RecognizeUtf8Text
            //ExSummary:Shows how to detect UTF-8 characters while loading an RTF document.
            // Create an "RtfLoadOptions" object to modify how we load an RTF document.
            RtfLoadOptions loadOptions = new RtfLoadOptions();

            // Set the "RecognizeUtf8Text" property to "false" to assume that the document uses the ISO 8859-1 charset
            // and loads every character in the document.
            // Set the "RecognizeUtf8Text" property to "true" to parse any variable-length characters that may occur in the text.
            loadOptions.RecognizeUtf8Text = recognizeUtf8Text;

            Document doc = new Document(MyDir + "UTF-8 characters.rtf", loadOptions);

            Assert.AreEqual(
                recognizeUtf8Text
                    ? "“John Doe´s list of currency symbols”™\r" +
                      "€, ¢, £, ¥, ¤"
                    : "â€œJohn DoeÂ´s list of currency symbolsâ€\u009dâ„¢\r" +
                      "â‚¬, Â¢, Â£, Â¥, Â¤",
                doc.FirstSection.Body.GetText().Trim());
            //ExEnd
        }
    }
}