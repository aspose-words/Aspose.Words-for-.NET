// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Markup;

using NUnit.Framework;

using System.IO;

namespace ApiExamples
{
    /// <summary>
    /// Tests that verify work with structured document tags in the document 
    /// </summary>
    [TestFixture]
    internal class ExStructuredDocumentTag : ApiExampleBase
    {
        [Test]
        public void RepeatingSection()
        {
            Document doc = new Document(MyDir + "TestRepeatingSection.docx");
            NodeCollection sdts = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);

            //Assert that the node have sdttype - RepeatingSection and it's not detected as RichText
            StructuredDocumentTag sdt = (StructuredDocumentTag)sdts[0];
            Assert.AreEqual(SdtType.RepeatingSection, sdt.SdtType);

            //Assert that the node have sdttype - RichText 
            sdt = (StructuredDocumentTag)sdts[1];
            Assert.AreNotEqual(SdtType.RepeatingSection, sdt.SdtType);
        }

        [Test]
        public void CheckBox()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            StructuredDocumentTag sdtCheckBox = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline);
            sdtCheckBox.Checked = true;

            //Insert content control into the document
            builder.InsertNode(sdtCheckBox);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            NodeCollection sdts = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
            
            StructuredDocumentTag sdt = (StructuredDocumentTag)sdts[0];
            Assert.AreEqual(true, sdt.Checked);
        }
    }
}
