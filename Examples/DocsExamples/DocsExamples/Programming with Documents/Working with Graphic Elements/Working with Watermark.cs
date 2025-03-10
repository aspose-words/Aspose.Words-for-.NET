using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Working_with_Graphic_Elements
{
    internal class WorkWithWatermark : DocsExamplesBase
    {
        [Test]
        public void AddTextWatermark()
        {
            //ExStart:AddTextWatermark
            //GistId:1f690a31c188a851d80d7aed4ff7e44c
            Document doc = new Document(MyDir + "Document.docx");

            TextWatermarkOptions options = new TextWatermarkOptions()
            {
                FontFamily = "Arial",
                FontSize = 36,
                Color = Color.Black,
                Layout = WatermarkLayout.Horizontal,
                IsSemitrasparent = false
            };

            doc.Watermark.SetText("Test", options);

            doc.Save(ArtifactsDir + "WorkWithWatermark.AddTextWatermark.docx");
            //ExEnd:AddTextWatermark
        }

#if NET48
        [Test]
        public void AddImageWatermark()
        {
            //ExStart:AddImageWatermark
            //GistId:1f690a31c188a851d80d7aed4ff7e44c
            Document doc = new Document(MyDir + "Document.docx");

            ImageWatermarkOptions options = new ImageWatermarkOptions
            {
                Scale = 5,
                IsWashout = false
            };

            doc.Watermark.SetImage(Image.FromFile(ImagesDir + "Transparent background logo.png"), options);

            doc.Save(ArtifactsDir + "WorkWithWatermark.AddImageWatermark.docx");
            //ExEnd:AddImageWatermark
        }

        [Test]
        public void RemoveDocumentWatermark()
        {
            //ExStart:RemoveDocumentWatermark
            //GistId:1f690a31c188a851d80d7aed4ff7e44c
            Document doc = new Document();

            // Add a plain text watermark.
            doc.Watermark.SetText("Aspose Watermark");

            // If we wish to edit the text formatting using it as a watermark,
            // we can do so by passing a TextWatermarkOptions object when creating the watermark.
            TextWatermarkOptions textWatermarkOptions = new TextWatermarkOptions();
            textWatermarkOptions.FontFamily = "Arial";
            textWatermarkOptions.FontSize = 36;
            textWatermarkOptions.Color = Color.Black;
            textWatermarkOptions.Layout = WatermarkLayout.Diagonal;
            textWatermarkOptions.IsSemitrasparent = false;

            doc.Watermark.SetText("Aspose Watermark", textWatermarkOptions);

            doc.Save(ArtifactsDir + "Document.TextWatermark.docx");

            // We can remove a watermark from a document like this.
            if (doc.Watermark.Type == WatermarkType.Text)
                doc.Watermark.Remove();

            doc.Save(ArtifactsDir + "WorkWithWatermark.RemoveDocumentWatermark.docx");
            //ExEnd:RemoveDocumentWatermark
        }
#endif

        //ExStart:AddDocumentWatermark
        //GistId:1f690a31c188a851d80d7aed4ff7e44c
        [Test]
        public void AddAndRemoveWatermark()
        {
            Document doc = new Document(MyDir + "Document.docx");

            InsertWatermarkText(doc, "CONFIDENTIAL");
            doc.Save(ArtifactsDir + "WorkWithWatermark.AddWatermark.docx");

            RemoveWatermarkShape(doc);
            doc.Save(ArtifactsDir + "WorkWithWatermark.RemoveWatermark.docx");
        }

        /// <summary>
        /// Inserts a watermark into a document.
        /// </summary>
        /// <param name="doc">The input document.</param>
        /// <param name="watermarkText">Text of the watermark.</param>
        private void InsertWatermarkText(Document doc, string watermarkText)
        {
            //ExStart:SetShapeName
            //GistId:1f690a31c188a851d80d7aed4ff7e44c
            // Create a watermark shape, this will be a WordArt shape.
            Shape watermark = new Shape(doc, ShapeType.TextPlainText) { Name = "Watermark" };
            //ExEnd:SetShapeName

            watermark.TextPath.Text = watermarkText;
            watermark.TextPath.FontFamily = "Arial";
            watermark.Width = 500;
            watermark.Height = 100;

            // Text will be directed from the bottom-left to the top-right corner.
            watermark.Rotation = -40;

            // Remove the following two lines if you need a solid black text.
            watermark.FillColor = Color.Gray; 
            watermark.StrokeColor = Color.Gray;

            // Place the watermark in the page center.
            watermark.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            watermark.RelativeVerticalPosition = RelativeVerticalPosition.Page;
            watermark.WrapType = WrapType.None;
            watermark.VerticalAlignment = VerticalAlignment.Center;
            watermark.HorizontalAlignment = HorizontalAlignment.Center;

            // Create a new paragraph and append the watermark to this paragraph.
            Paragraph watermarkPara = new Paragraph(doc);
            watermarkPara.AppendChild(watermark);

            // Insert the watermark into all headers of each document section.
            foreach (Section sect in doc.Sections)
            {
                // There could be up to three different headers in each section.
                // Since we want the watermark to appear on all pages, insert it into all headers.
                InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderPrimary);
                InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderFirst);
                InsertWatermarkIntoHeader(watermarkPara, sect, HeaderFooterType.HeaderEven);
            }
        }

        private void InsertWatermarkIntoHeader(Paragraph watermarkPara, Section sect,
            HeaderFooterType headerType)
        {
            HeaderFooter header = sect.HeadersFooters[headerType];

            if (header == null)
            {
                // There is no header of the specified type in the current section, so we need to create it.
                header = new HeaderFooter(sect.Document, headerType);
                sect.HeadersFooters.Add(header);
            }

            // Insert a clone of the watermark into the header.
            header.AppendChild(watermarkPara.Clone(true));
        }
        //ExEnd:AddDocumentWatermark

        //ExStart:RemoveWatermarkShape
        //GistId:1f690a31c188a851d80d7aed4ff7e44c
        private void RemoveWatermarkShape(Document doc)
        {
            foreach (HeaderFooter hf in doc.GetChildNodes(NodeType.HeaderFooter, true))
            {
                foreach (Shape shape in hf.GetChildNodes(NodeType.Shape, true))
                {
                    if (shape.Name.Contains("Watermark"))
                    {
                        shape.Remove();
                    }
                }
            }
        }
        //ExEnd:RemoveWatermarkShape
    }
}
