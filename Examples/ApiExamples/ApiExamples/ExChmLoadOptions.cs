using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    public class ExChmLoadOptions : ApiExampleBase
    {
        [Test] //Need to add tests
        public void OriginalFileName()
        {
            //ExStart
            //ExFor:ChmLoadOptions.OriginalFileName
            //ExSummary:Shows how to resolve URLs like "ms-its:myfile.chm::/index.htm".
            // Our document contains URLs like "ms-its:amhelp.chm::....htm", but it has a different name,
            // so file links don't work after saving it to HTML.
            // We need to define the original filename in 'ChmLoadOptions' to avoid this behavior.
            ChmLoadOptions loadOptions = new ChmLoadOptions { OriginalFileName = "amhelp.chm" };

            Document doc = new Document(new MemoryStream(File.ReadAllBytes(MyDir + "Document with ms-its links.chm")),
                loadOptions);
            
            doc.Save(ArtifactsDir + "ExChmLoadOptions.OriginalFileName.html");
            //ExEnd
        }
    }
}
