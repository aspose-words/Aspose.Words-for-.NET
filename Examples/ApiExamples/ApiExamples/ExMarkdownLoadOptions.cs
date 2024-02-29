using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Loading;
using NUnit.Framework;

namespace ApiExamples
{
    class ExMarkdownLoadOptions : ApiExampleBase
    {
        [Test]
        public void PreserveEmptyLines()
        {
            //ExStart:PreserveEmptyLines
            //GistId:a775441ecb396eea917a2717cb9e8f8f
            //ExFor:MarkdownLoadOptions
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
    }
}
