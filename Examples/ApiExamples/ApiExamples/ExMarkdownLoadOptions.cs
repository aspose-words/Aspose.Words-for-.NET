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

                Assert.AreEqual("\rLine1\r\rLine2\r\f", doc.GetText());
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
                Assert.AreEqual(Underline.Single, para.Runs[0].Font.Underline);

                loadOptions = new MarkdownLoadOptions() { ImportUnderlineFormatting = false };
                doc = new Document(stream, loadOptions);

                para = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
                Assert.AreEqual(Underline.None, para.Runs[0].Font.Underline);
            }
            //ExEnd:ImportUnderlineFormatting
        }
    }
}
