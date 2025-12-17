// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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
    public class ExComHelper : ApiExampleBase
    {
        [Test]
        public void ComHelper()
        {
            //ExStart
            //ExFor:ComHelper
            //ExFor:ComHelper.#ctor
            //ExFor:ComHelper.Open(Stream)
            //ExFor:ComHelper.Open(String)
            //ExSummary:Shows how to open documents using the ComHelper class.
            // The ComHelper class allows us to load documents from within COM clients.
            ComHelper comHelper = new ComHelper();

            // 1 -  Using a local system filename:
            Document doc = comHelper.Open(MyDir + "Document.docx");

            Assert.That(doc.GetText().Trim(), Is.EqualTo("Hello World!\r\rHello Word!\r\r\rHello World!"));

            // 2 -  From a stream:
            using (FileStream stream = new FileStream(MyDir + "Document.docx", FileMode.Open))
            {
                doc = comHelper.Open(stream);

                Assert.That(doc.GetText().Trim(), Is.EqualTo("Hello World!\r\rHello Word!\r\r\rHello World!"));
            }
            //ExEnd
        }
    }
}
