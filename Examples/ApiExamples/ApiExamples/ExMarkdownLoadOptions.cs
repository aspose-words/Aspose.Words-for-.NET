// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExMarkdownLoadOptions : ApiExampleBase
    {
        [Test]
        public void PreserveEmptyLines()
        {
            //ExStart:PreserveEmptyLines
            //GistId:a775441ecb396eea917a2717cb9e8f8f
            //ExFor:MarkdownLoadOptions
            //ExFor:MarkdownLoadOptions.#ctor
            //ExFor:MarkdownLoadOptions.PreserveEmptyLines
            //ExSummary:Shows how to preserve empty line while load a document.
            string mdText = $"{Environment.NewLine}Line1{Environment.NewLine}{Environment.NewLine}Line2{Environment.NewLine}{Environment.NewLine}";
            using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(mdText)))
            {
                MarkdownLoadOptions loadOptions = new MarkdownLoadOptions() { PreserveEmptyLines = true };
                Document doc = new Document(stream, loadOptions);

                Assert.That(doc.GetText(), Is.EqualTo("\rLine1\r\rLine2\r\f"));
            }
            //ExEnd:PreserveEmptyLines
        }

        [Test]
        public void ImportUnderlineFormatting()
        {
            //ExStart:ImportUnderlineFormatting
            //GistId:e06aa7a168b57907a5598e823a22bf0a
            //ExFor:MarkdownLoadOptions.ImportUnderlineFormatting
            //ExSummary:Shows how to recognize plus characters "++" as underline text formatting.
            using (MemoryStream stream = new MemoryStream(Encoding.ASCII.GetBytes("++12 and B++")))
            {
                MarkdownLoadOptions loadOptions = new MarkdownLoadOptions() { ImportUnderlineFormatting = true };
                Document doc = new Document(stream, loadOptions);

                Paragraph para = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
                Assert.That(para.Runs[0].Font.Underline, Is.EqualTo(Underline.Single));

                loadOptions = new MarkdownLoadOptions() { ImportUnderlineFormatting = false };
                doc = new Document(stream, loadOptions);

                para = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
                Assert.That(para.Runs[0].Font.Underline, Is.EqualTo(Underline.None));
            }
            //ExEnd:ImportUnderlineFormatting
        }

        [Test]
        public void SoftLineBreakCharacter()
        {
            //ExStart:SoftLineBreakCharacter
            //GistId:571cc6e23284a2ec075d15d4c32e3bbf
            //ExFor:MarkdownLoadOptions.SoftLineBreakCharacter
            //ExSummary:Shows how to set soft line break character.
            using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes("line1\nline2")))
            {
                MarkdownLoadOptions loadOptions = new MarkdownLoadOptions();
                loadOptions.SoftLineBreakCharacter = ControlChar.LineBreakChar;
                Document doc = new Document(stream, loadOptions);

                Assert.That(doc.GetText().Trim(), Is.EqualTo("line1\u000bline2"));
            }
            //ExEnd:SoftLineBreakCharacter
        }
    }
}
