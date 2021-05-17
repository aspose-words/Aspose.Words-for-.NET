using Aspose.Words;
using Aspose.Words.Loading;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Load_Options
{
    public class WorkingWithRtfLoadOptions : DocsExamplesBase
    {
        [Test]
        public void RecognizeUtf8Text()
        {
            //ExStart:RecognizeUtf8Text
            RtfLoadOptions loadOptions = new RtfLoadOptions { RecognizeUtf8Text = true };

            Document doc = new Document(MyDir + "UTF-8 characters.rtf", loadOptions);

            doc.Save(ArtifactsDir + "WorkingWithRtfLoadOptions.RecognizeUtf8Text.rtf");
            //ExEnd:RecognizeUtf8Text
        }
    }
}