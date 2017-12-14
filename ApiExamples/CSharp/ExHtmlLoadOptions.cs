// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
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
    internal class ExHtmlLoadOptions : ApiExampleBase
    {
        [Test]
        public void SupportVml()
        {
            //ExStart
            //ExFor:HtmlLoadOptions.SupportVml
            //ExSummary:Shows how to parse HTML document with conditional comments like "<!--[if gte vml 1]>" and "<![if !vml]>"
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            //If value is true, then we parse "<!--[if gte vml 1]>", else parse "<![if !vml]>"
            loadOptions.SupportVml = true;
            //Wait for a response, when loading external resources
            loadOptions.WebRequestTimeout = 1000;

            Document doc = new Document(MyDir + "Shape.VmlAndDml.htm", loadOptions);
            doc.Save(MyDir + @"\Artifacts\Shape.VmlAndDml.docx");
            //ExEnd
        }

        //This is just a test, no need adding example tags.
        [Test]
        public void WebRequestTimeoutDefaultValue()
        {
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Assert.AreEqual(100000, loadOptions.WebRequestTimeout);
        }
    }
}