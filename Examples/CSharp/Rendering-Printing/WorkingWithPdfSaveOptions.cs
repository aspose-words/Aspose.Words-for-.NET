using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class WorkingWithPdfSaveOptions
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting();

            EscapeUriInPdf(dataDir);
            ExportHeaderFooterBookmarks(dataDir);
            ScaleWmfFontsToMetafileSize(dataDir);
            AdditionalTextPositioning(dataDir);
            ConversionToPDF17(dataDir);
            DownsamplingImages(dataDir);
            SaveToPdfWithOutline(dataDir);
            CustomPropertiesExport(dataDir);
            //ExportDocumentStructure(dataDir);
            //PdfImageComppression(dataDir);
            UpdateIfLastPrinted(dataDir);
            EffectsRendering(dataDir);
        }

        public static void EscapeUriInPdf(String dataDir)
        {
            // ExStart:EscapeUriInPdf
            // The path to the documents directory.
            Document doc = new Document(dataDir + "EscapeUri.docx");

            PdfSaveOptions options = new PdfSaveOptions();
            options.EscapeUri = false;

            dataDir = dataDir + "EscapeUri_out.pdf";
            doc.Save(dataDir, options);
            // ExEnd:EscapeUriInPdf
            Console.WriteLine("\nFile saved at " + dataDir);
        }

        public static void ExportHeaderFooterBookmarks(String dataDir)
        {
            // ExStart:ExportHeaderFooterBookmarks
            // The path to the documents directory.
            Document doc = new Document(dataDir + "TestFile.docx");

            PdfSaveOptions options = new PdfSaveOptions();
            options.OutlineOptions.DefaultBookmarksOutlineLevel = 1;
            options.HeaderFooterBookmarksExportMode = HeaderFooterBookmarksExportMode.First;

            dataDir = dataDir + "ExportHeaderFooterBookmarks_out.pdf";
            doc.Save(dataDir, options);
            // ExEnd:ExportHeaderFooterBookmarks
            Console.WriteLine("\nFile saved at " + dataDir);
        }

        public static void ScaleWmfFontsToMetafileSize(String dataDir)
        {
            // ExStart:ScaleWmfFontsToMetafileSize
            // The path to the documents directory.
            Document doc = new Document(dataDir + "MetafileRendering.docx");

            MetafileRenderingOptions metafileRenderingOptions =
                       new MetafileRenderingOptions
                       {
                           ScaleWmfFontsToMetafileSize = false
                       };

            // If Aspose.Words cannot correctly render some of the metafile records to vector graphics then Aspose.Words renders this metafile to a bitmap. 
            PdfSaveOptions options = new PdfSaveOptions { MetafileRenderingOptions = metafileRenderingOptions };

            dataDir = dataDir + "ScaleWmfFontsToMetafileSize_out.pdf";
            doc.Save(dataDir, options);
            // ExEnd:ScaleWmfFontsToMetafileSize
            Console.WriteLine("\nFonts as metafile are rendered to its default size in PDF. File saved at " + dataDir);
        }
        
        public static void AdditionalTextPositioning(string dataDir)
        {
            // ExStart:AdditionalTextPositioning
            // The path to the documents directory.
            Document doc = new Document(dataDir + "TestFile.docx");

            PdfSaveOptions options = new PdfSaveOptions();
            options.AdditionalTextPositioning = true;

            dataDir = dataDir + "AdditionalTextPositioning_out.pdf";
            doc.Save(dataDir, options);
            // ExEnd:AdditionalTextPositioning
            Console.WriteLine("\nFile saved at " + dataDir);
        }

        public static void ConversionToPDF17(string dataDir)
        {
            // ExStart:ConversionToPDF17
            // The path to the documents directory.
            Document originalDoc = new Document(dataDir + "Rendering.doc");

            // Provide PDFSaveOption compliance to PDF17
            // or just convert without SaveOptions
            PdfSaveOptions pso = new PdfSaveOptions();
            pso.Compliance = PdfCompliance.Pdf17;

            originalDoc.Save(dataDir + "Output.pdf", pso);
            // ExEnd:ConversionToPDF17
            Console.WriteLine("\nFile saved at " + dataDir);
        }

        public static void DownsamplingImages(string dataDir)
        {
            // ExStart:DownsamplingImages
            // Open a document that contains images 
            Document doc = new Document(dataDir + "Rendering.doc");

            // If we want to convert the document to .pdf, we can use a SaveOptions implementation to customize the saving process
            PdfSaveOptions options = new PdfSaveOptions();

            // We can set the output resolution to a different value
            // The first two images in the input document will be affected by this
            options.DownsampleOptions.Resolution = 36;

            // We can set a minimum threshold for downsampling 
            // This value will prevent the second image in the input document from being downsampled
            options.DownsampleOptions.ResolutionThreshold = 128;

            doc.Save(dataDir + "PdfSaveOptions.DownsampleOptions.pdf", options);
            // ExEnd:DownsamplingImages
        }

        public static void SaveToPdfWithOutline(string dataDir)
        {
            // ExStart:SaveToPdfWithOutline
            // Open a document
            Document doc = new Document(dataDir + "Rendering.doc");

            PdfSaveOptions options = new PdfSaveOptions();
            options.OutlineOptions.HeadingsOutlineLevels = 3;
            options.OutlineOptions.ExpandedOutlineLevels = 1;

            doc.Save(dataDir + "Rendering.SaveToPdfWithOutline.pdf", options);
            // ExEnd:SaveToPdfWithOutline
        }

        public static void CustomPropertiesExport(string dataDir)
        {
            // ExStart:CustomPropertiesExport
            // Open a document
            Document doc = new Document();

            // Add a custom document property that doesn't use the name of some built in properties
            doc.CustomDocumentProperties.Add("Company", "My value");

            // Configure the PdfSaveOptions like this will display the properties
            // in the "Document Properties" menu of Adobe Acrobat Pro
            PdfSaveOptions options = new PdfSaveOptions();
            options.CustomPropertiesExport = PdfCustomPropertiesExport.Standard;

            doc.Save(dataDir + "PdfSaveOptions.CustomPropertiesExport.pdf", options);
            // ExEnd:CustomPropertiesExport
        }

        public static void ExportDocumentStructure(string dataDir)
        {
            // ExStart:ExportDocumentStructure
            // Open a document
            Document doc = new Document(dataDir + "Paragraphs.docx");

            // Create a PdfSaveOptions object and configure it to preserve the logical structure that's in the input document
            // The file size will be increased and the structure will be visible in the "Content" navigation pane
            // of Adobe Acrobat Pro, while editing the .pdf
            PdfSaveOptions options = new PdfSaveOptions();
            options.ExportDocumentStructure = true;

            doc.Save(dataDir + "PdfSaveOptions.ExportDocumentStructure.pdf", options);
            // ExEnd:ExportDocumentStructure
        }

        public static void PdfImageComppression(string dataDir)
        {
            // ExStart:PdfImageComppression
            // Open a document
            Document doc = new Document(dataDir + "SaveOptions.PdfImageCompression.rtf");

            PdfSaveOptions options = new PdfSaveOptions
            {
                ImageCompression = PdfImageCompression.Jpeg,
                PreserveFormFields = true
            };
            
            doc.Save(dataDir + "SaveOptions.PdfImageCompression.pdf", options);

            PdfSaveOptions optionsA1B = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1b,
                ImageCompression = PdfImageCompression.Jpeg,

                // Use JPEG compression at 50% quality to reduce file size
                JpegQuality = 100, 
                ImageColorSpaceExportMode = PdfImageColorSpaceExportMode.SimpleCmyk
            };
            
            doc.Save(dataDir + "SaveOptions.PdfImageComppression PDF_A_1_B.pdf", optionsA1B);
            // ExEnd:PdfImageComppression
            Console.WriteLine("\nFile saved at " + dataDir);
        }

        public static void UpdateIfLastPrinted(string dataDir)
        {
            // ExStart:UpdateIfLastPrinted
            // Open a document
            Document doc = new Document(dataDir + "Rendering.doc");

            SaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.UpdateLastPrintedProperty = false;

            doc.Save(dataDir + "PdfSaveOptions.UpdateIfLastPrinted.pdf", saveOptions);
            // ExEnd:UpdateIfLastPrinted
        }

        public static void EffectsRendering(string dataDir)
        {
            // ExStart:EffectsRendering
            // Open a document
            Document doc = new Document(dataDir + "Rendering.doc");

            SaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.Dml3DEffectsRenderingMode = Dml3DEffectsRenderingMode.Advanced;
            
            doc.Save(dataDir, saveOptions);
            // ExEnd:EffectsRendering
        }

        public static void SetImageInterpolation(string dataDir)
        {
            // ExStart:SetImageInterpolation
            Document doc = new Document(dataDir);

            PdfSaveOptions saveOptions = new PdfSaveOptions();
            saveOptions.InterpolateImages = true;
            
            doc.Save(dataDir, saveOptions);
            // ExEnd:SetImageInterpolation
        }
    }
}
