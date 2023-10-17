using Aspose.Words;
using NUnit.Framework;

namespace PluginsExamples
{
    public class ConverterPlugin : PluginsExamplesBase
    {
        [Test]
        public void ConvertDocument()
        {
            //ExStart:ConvertDocument
            //GistId:8dbc9dbdbd4cb3e3bccb77d95f64d88a
            var doc = new Document(MyDir + "Document.docx");

            doc.Save(ArtifactsDir + "ConverterPlugin.ConvertDocument.pdf");
            //ExEnd:ConvertDocument
        }
    }
}
