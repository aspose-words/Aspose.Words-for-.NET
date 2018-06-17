﻿// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExDigitalSignatureCollection : ApiExampleBase
    {
        [Test]
        public void GetEnumerator()
        {
            //ExStart
            //ExFor:DigitalSignatureCollection.GetEnumerator
            //ExSummary:Shows how to load and enumerate all digital signatures of a document.
            DigitalSignatureCollection digitalSignatures = DigitalSignatureUtil.LoadSignatures(MyDir + "Document.DigitalSignature.docx");

            IEnumerator enumerator = digitalSignatures.GetEnumerator();
            while (enumerator.MoveNext())
            {
                // Do something useful
                DigitalSignature ds = (DigitalSignature)enumerator.Current;

                if (ds != null)
                    Console.WriteLine(ds.ToString());
            }
            //ExEnd
        }
    }
}