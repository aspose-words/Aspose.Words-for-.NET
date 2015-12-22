// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.Document
{
    [TestFixture]
    public class ExDigitalSignatureUtil : QaTestsBase
    {
        [Test]
        public void RemoveAllSignaturesEx()
        {
            //ExStart
            //ExFor:RemoveAllSignatures
            //ExId:RemoveAllSignaturesEx
            //ExSummary:Shows how to use RemoveAllSignatures.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");

            string outputDocFileName = MyDir + "Document.NoSignatures.doc";
            Aspose.Words.DigitalSignatureUtil.RemoveAllSignatures(doc.OriginalFileName, outputDocFileName);            
            //ExEnd
        }
    }
}
