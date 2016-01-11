// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.DigitalSignature
{
    [TestFixture]
    public class ExDigitalSignature : QaTestsBase
    {
        [Test]
        public void ToStringEx()
        {
            //ExStart
            //ExFor:ToString
            //ExId:ToStringEx
            //ExSummary:Shows how to use ToString.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.Signed.docx");
            Console.WriteLine(doc.DigitalSignatures[0]);         
            //ExEnd
        }
    }
}
