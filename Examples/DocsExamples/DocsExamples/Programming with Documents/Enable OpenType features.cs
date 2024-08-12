﻿using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents
{
    internal class EnableOpenTypeFeatures : DocsExamplesBase
    {
        [Test]
        public void OpenTypeFeatures()
        {
            //ExStart:OpenTypeFeatures
            //GistId:7840fae2297fa05bba1ca0608cb81bf1
            Document doc = new Document(MyDir + "OpenType text shaping.docx");

            // When we set the text shaper factory, the layout starts to use OpenType features.
            // An Instance property returns BasicTextShaperCache object wrapping HarfBuzzTextShaperFactory.
            doc.LayoutOptions.TextShaperFactory = Aspose.Words.Shaping.HarfBuzz.HarfBuzzTextShaperFactory.Instance;

            doc.Save(ArtifactsDir + "WorkingWithHarfBuzz.OpenTypeFeatures.pdf");
            //ExEnd:OpenTypeFeatures
        }
    }
}