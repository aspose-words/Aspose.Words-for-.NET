using Aspose.Words;
using NUnit.Framework;
using System.Drawing;

namespace PluginsExamples
{
    public class WatermarkPlugin : PluginsExamplesBase
    {
        [Test]
        public void AddWatermark()
        {
            //ExStart:AddWatermark
            //GistId:f1cc2c41c92a748dda99a825cd08c616
            var doc = new Document(MyDir + "Document.docx");
            doc.Watermark.SetText("Watermark Text");
            doc.Save(ArtifactsDir + "WatermarkPlugin.AddWatermark.docx");
            //ExEnd:AddWatermark
        }

        [Test]
        public void AddWatermarkWithFormatting()
        {
            //ExStart:AddWatermarkWithFormatting
            //GistId:f1cc2c41c92a748dda99a825cd08c616
            var doc = new Document("Document.docx");

            // You can edit the text formatting using it as a watermark.
            var textWatermarkOptions = new TextWatermarkOptions();
            textWatermarkOptions.Color = Color.Black;
            textWatermarkOptions.Layout = WatermarkLayout.Diagonal;

            doc.Watermark.SetText("Watermark Text", textWatermarkOptions);
            doc.Save(ArtifactsDir + "WatermarkPlugin.AddWatermarkWithFormatting.docx");
            //ExEnd:AddWatermarkWithFormatting
        }

        [Test]
        public void RemoveWatermark()
        {
            //ExStart:RemoveWatermark
            //GistId:f1cc2c41c92a748dda99a825cd08c616
            var doc = new Document(MyDir + "Document.docx");

            if (doc.Watermark.Type == WatermarkType.Text)
                doc.Watermark.Remove();

            doc.Save(ArtifactsDir + "WatermarkPlugin.RemoveWatermark.docx");
            //ExEnd:RemoveWatermark
        }
    }
}
