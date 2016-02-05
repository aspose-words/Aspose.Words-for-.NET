// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using NUnit.Framework;


namespace ApiExamples.DigitalSignature
{
    [TestFixture]
    public class ExDigitalSignature : ApiExampleBase
    {
        [Test]
        public void ToStringEx()
        {
            //ExStart
            //ExFor:DigitalSignature.ToString
            //ExSummary:Shows how to get the string representation of a signature from a signed document.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.Signed.docx");
            Console.WriteLine(doc.DigitalSignatures[0].ToString());
            //ExEnd
        }
    }
}