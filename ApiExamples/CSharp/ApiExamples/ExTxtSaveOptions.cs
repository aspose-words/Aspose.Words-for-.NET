// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
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
    public class ExTxtSaveOptions : ApiExampleBase
    {
        [Test]
        [TestCase(
            "Some text before page break\r\rJidqwjidqwojidqwojidqwojidqwojidqwoji\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\rQwdqwdqwdqwdqwdqwdqwd\rQwdqwdqwdqwdqwdqwdqw\r\r\r\r\rqwdqwdqwdqwdqwdqwdqwqwd\r\f",
            false)]
        [TestCase(
            "Some text before page break\r\f\r\fJidqwjidqwojidqwojidqwojidqwojidqwoji\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\r\rQwdqwdqwdqwdqwdqwdqwd\rQwdqwdqwdqwdqwdqwdqw\r\r\r\r\f\r\fqwdqwdqwdqwdqwdqwdqwqwd\r\f",
            true)]
        public void PageBreaks(string resultText, bool isForcePageBreaks)
        {
            //ExStart
            //ExFor:TxtSaveOptions.ForcePageBreaks
            //ExSummary:Shows how to specify whether the page breaks should be preserved during export.
            Document doc = new Document(MyDir + "SaveOptions.PageBreaks.docx");

            TxtSaveOptions saveOptions = new TxtSaveOptions
            {
                ForcePageBreaks = isForcePageBreaks
            };

            doc.Save(MyDir + @"\Artifacts\SaveOptions.PageBreaks.txt", saveOptions);
            //ExEnd

            Document docFalse = new Document(MyDir + @"\Artifacts\SaveOptions.PageBreaks.txt");
            Assert.AreEqual(resultText, docFalse.GetText());
        }
    }
}