using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Working_with_Graphic_Elements
{
    internal class WorkingWithBarcodeGenerator : DocsExamplesBase
    {
        [Test]
        public void BarcodeGenerator()
        {
            //ExStart:BarcodeGenerator
            //GistId:00d34dba66626dbc0175b60bb3b71c8a
            Document doc = new Document(MyDir + "Field sample - BARCODE.docx");

            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();
            
            doc.Save(ArtifactsDir + "WorkingWithBarcodeGenerator.BarcodeGenerator.pdf");
            //ExEnd:BarcodeGenerator
        }
    }
}