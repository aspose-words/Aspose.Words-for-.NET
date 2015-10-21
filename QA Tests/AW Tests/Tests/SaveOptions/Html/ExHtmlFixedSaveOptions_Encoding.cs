using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace QA_Tests.Tests.SaveOptions.Html
{
    [TestFixture]
    internal class ExHtmlFixedSaveOptionsEncoding : QaTestsBase
    {
        [Test]
        public void UseEncoding()
        {
            //ExStart
            //ExFor:Saving
            //ExFor:Saving.HtmlFixedSaveOptions
            //ExSummary:Shows how to use "Encoding" parameter with "HtmlFixedSaveOptions"
            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello World!");

            //Create "HtmlFixedSaveOptions" with "Encoding" parameter
            //You can also set "Encoding" using System.Text.Encoding, like "Encoding.ASCII", or "Encoding.GetEncoding()"
            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                Encoding = new ASCIIEncoding(),
                SaveFormat = SaveFormat.HtmlFixed,
            };

            //Uses "HtmlFixedSaveOptions"
            doc.Save(MyDir + "UseEncoding.html", htmlFixedSaveOptions);
            //ExEnd
        }
    }
}
