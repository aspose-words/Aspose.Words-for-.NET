﻿using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class SaveDocWithHtmlSaveOptions
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();
            SaveHtmlWithMetafileFormat(dataDir);
            ImportExportSVGinHTML(dataDir);
            SetCssClassNamePrefix(dataDir);
            SetExportCidUrlsForMhtmlResources(dataDir);
        }

        public static void SaveHtmlWithMetafileFormat(string dataDir)
        {
            // ExStart:SaveHtmlWithMetafileFormat
            Document doc = new Document(dataDir + "Document.docx");
            HtmlSaveOptions options = new HtmlSaveOptions();
            options.MetafileFormat = HtmlMetafileFormat.EmfOrWmf;

            dataDir = dataDir + "SaveHtmlWithMetafileFormat_out.html";
            doc.Save(dataDir, options);
            // ExEnd:SaveHtmlWithMetafileFormat
            Console.WriteLine("\nDocument saved with Metafile format.\nFile saved at " + dataDir);
        }

        public static void ImportExportSVGinHTML(string dataDir)
        {
            // ExStart:ImportExportSVGinHTML
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Here is an SVG image: ");
            builder.InsertHtml(
                @"<svg height='210' width='500'>
                <polygon points='100,10 40,198 190,78 10,78 160,198' 
                    style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />
            </svg> ");

            HtmlSaveOptions options = new HtmlSaveOptions();
            options.MetafileFormat = HtmlMetafileFormat.Svg;

            dataDir = dataDir + "ExportSVGinHTML_out.html";
            doc.Save(dataDir, options);
            // ExEnd:ImportExportSVGinHTML
            Console.WriteLine("\nDocument saved with SVG Metafile format.\nFile saved at " + dataDir);
        }

        public static void SetCssClassNamePrefix(string dataDir)
        {
            // ExStart:SetCssClassNamePrefix
            Document doc = new Document(dataDir + "Document.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.CssStyleSheetType = CssStyleSheetType.External;
            saveOptions.CssClassNamePrefix = "pfx_";

            dataDir = dataDir + "CssClassNamePrefix_out.html";
            doc.Save(dataDir, saveOptions);
            // ExEnd:SetCssClassNamePrefix
            Console.WriteLine("\nDocument saved with CSS prefix pfx_.\nFile saved at " + dataDir);
        }

        public static void SetExportCidUrlsForMhtmlResources(string dataDir)
        {
            // ExStart:SetExportCidUrlsForMhtmlResources
            var doc = new Aspose.Words.Document(dataDir + "CidUrls.docx");

            Aspose.Words.Saving.HtmlSaveOptions saveOptions = new Aspose.Words.Saving.HtmlSaveOptions(SaveFormat.Mhtml);
            saveOptions.PrettyFormat = true;
            saveOptions.ExportCidUrlsForMhtmlResources = true;
            saveOptions.SaveFormat = SaveFormat.Mhtml;

            dataDir = dataDir + "SetExportCidUrlsForMhtmlResources_out.mhtml";
            doc.Save(dataDir, saveOptions);
            // ExEnd:SetExportCidUrlsForMhtmlResources
            Console.WriteLine("\nDocument has saved with Content - Id URL scheme.\nFile saved at " + dataDir);
        }
    }
}
