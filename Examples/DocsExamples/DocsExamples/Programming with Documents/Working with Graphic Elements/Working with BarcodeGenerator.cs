using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Working_with_Graphic_Elements
{
    internal class WorkingWithBarcodeGenerator : DocsExamplesBase
    {
        [Test]
        public void GenerateACustomBarCodeImage()
        {
            //ExStart:GenerateACustomBarCodeImage
            Document doc = new Document(MyDir + "Field sample - BARCODE.docx");

            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();
            
            doc.Save(ArtifactsDir + "WorkingWithBarcodeGenerator.GenerateACustomBarCodeImage.pdf");
            //ExEnd:GenerateACustomBarCodeImage
        }
    }
}