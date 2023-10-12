using Aspose.Words;
using NUnit.Framework;

namespace PluginsExamples
{
    public class ProcessorChmPlugin : PluginsExamplesBase
    {
        [Test]
        public void ReadChm()
        {
            //ExStart:ReadChm
            //GistId:cae9721622323e1d3c3812a30687c3af
            var doc = new Document(MyDir + "HTML help.chm");
            // We provide only reading .chm document, it's up to developer
            // to choose the library for further saving document.
            doc.Save(ArtifactsDir + "ProcessorChmPlugin.ReadChm.docx");
            //ExEnd:ReadChm
        }
    }
}
