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
        //ToDo: Add gold asserts
        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void SupportVml(bool supportVml)
        {
            //ExStart
            //ExFor:HtmlLoadOptions.SupportVml
            //ExSummary:Demonstrates how to parse html document with conditional comments like "&lt;!--[if gte vml 1]&gt;" and "&lt;![if !vml]&gt;"
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();

            //If SupportVml = true, then we parse "&lt;!--[if gte vml 1]&gt;", else parse "&lt;![if !vml]&gt;"
            loadOptions.SupportVml = supportVml;
            //Wait for a response, when loading external resources
            loadOptions.WebRequestTimeout = 1000;

            Document doc = new Document(MyDir + "Shape.VmlAndDml.htm", loadOptions);
            doc.Save(MyDir + @"\Artifacts\Shape.VmlAndDml.docx");
            //ExEnd
        }

        [Test]
        public void WebRequestTimeoutDefaultValue()
        {
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Assert.AreEqual(100000, loadOptions.WebRequestTimeout);
        }
    }
}