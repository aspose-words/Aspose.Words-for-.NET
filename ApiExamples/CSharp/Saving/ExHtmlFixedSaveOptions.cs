using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;


namespace ApiExamples.Saving
{
    [TestFixture]
    internal class ExHtmlFixedSaveOptions : ApiExampleBase
    {
        [Test]
        public void UseEncoding()
        {
            //ExStart
            //ExFor:Saving.HtmlFixedSaveOptions.Encoding
            //ExSummary:Shows how to use "Encoding" parameter with "HtmlFixedSaveOptions"
            Aspose.Words.Document doc = new Aspose.Words.Document();

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
