using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Save_Options
{
    public class WorkingWithOdtSaveOptions : DocsExamplesBase
    {
        [Test]
        public void MeasureUnit()
        {
            //ExStart:MeasureUnit
            Document doc = new Document(MyDir + "Document.docx");

            // Open Office uses centimeters when specifying lengths, widths and other measurable formatting
            // and content properties in documents whereas MS Office uses inches.
            OdtSaveOptions saveOptions = new OdtSaveOptions { MeasureUnit = OdtSaveMeasureUnit.Inches };

            doc.Save(ArtifactsDir + "WorkingWithOdtSaveOptions.MeasureUnit.odt", saveOptions);
            //ExEnd:MeasureUnit
        }
    }
}