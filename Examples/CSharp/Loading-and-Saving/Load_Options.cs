using Aspose.Words.Markup;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words.Loading;

namespace Aspose.Words.Examples.CSharp.Loading_and_Saving
{
    class Load_Options
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_QuickStart();

            LoadOptionsUpdateDirtyFields(dataDir);
            LoadAndSaveEncryptedODT(dataDir);
            VerifyODTdocument(dataDir);
            ConvertShapeToOfficeMath(dataDir);
            SetMSWordVersion(dataDir);
            LoadOptionsWarningCallback(dataDir);
            LoadOptionsResourceLoadingCallback(dataDir);
            LoadOptionsEncoding(dataDir);
            SkipPdfImages(dataDir);
            ConvertMetafilesToPng(dataDir);
        }

        public static void LoadOptionsUpdateDirtyFields(string dataDir)
        {
            // ExStart:LoadOptionsUpdateDirtyFields  
            LoadOptions lo = new LoadOptions();

            //Update the fields with the dirty attribute
            lo.UpdateDirtyFields = true;

            //Load the Word document
            Document doc = new Document(dataDir + @"input.docx", lo);

            //Save the document into DOCX
            doc.Save(dataDir + "output.docx", SaveFormat.Docx);
            // ExEnd:LoadOptionsUpdateDirtyFields 
            Console.WriteLine("\nUpdate the fields with the dirty attribute successfully.\nFile saved at " + dataDir);
        }

        public static void LoadAndSaveEncryptedODT(string dataDir)
        {
            // ExStart:LoadAndSaveEncryptedODT  
            Document doc = new Document(dataDir + @"encrypted.odt", new Aspose.Words.LoadOptions("password"));

            doc.Save(dataDir + "out.odt", new OdtSaveOptions("newpassword"));
            // ExEnd:LoadAndSaveEncryptedODT 
            Console.WriteLine("\nLoad and save encrypted document successfully.\nFile saved at " + dataDir);
        }

        public static void VerifyODTdocument(string dataDir)
        {
            // ExStart:VerifyODTdocument  
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(dataDir + @"encrypted.odt");
            Console.WriteLine(info.IsEncrypted);
            // ExEnd:VerifyODTdocument 
        }

        public static void ConvertShapeToOfficeMath(string dataDir)
        {
            // ExStart:ConvertShapeToOfficeMath   
            LoadOptions lo = new LoadOptions();
            lo.ConvertShapeToOfficeMath = true;

            // Specify load option to use previous default behaviour i.e. convert math shapes to office math ojects on loading stage.
            Document doc = new Document(dataDir + @"OfficeMath.docx", lo);
            //Save the document into DOCX
            doc.Save(dataDir + "ConvertShapeToOfficeMath_out.docx", SaveFormat.Docx);
            // ExEnd:ConvertShapeToOfficeMath  
        }

        public static void SetMSWordVersion(string dataDir)
        {
            // ExStart:SetMSWordVersion  
            // Create a new LoadOptions object, which will load documents according to MS Word 2019 specification by default.
            LoadOptions loadOptions = new LoadOptions();
            // Change the loading version to Microsoft Word 2010.
            loadOptions.MswVersion = MsWordVersion.Word2010;
            
            Document doc = new Document(dataDir + "document.docx", loadOptions);
            doc.Save(dataDir + "Word2003_out.docx");
            // ExEnd:SetMSWordVersion 
            Console.WriteLine("\n Loaded with MS Word Version successfully.\nFile saved at " + dataDir); 
        }

        public static void SetTempFolder(string dataDir)
        {
            // ExStart:SetTempFolder  
            LoadOptions lo = new LoadOptions();
            lo.TempFolder = @"C:\TempFolder\";

            Document doc = new Document(dataDir + "document.docx", lo);
            // ExEnd:SetTempFolder  
        }
        
        public static void LoadOptionsWarningCallback(string dataDir)
        {
            //ExStart:LoadOptionsWarningCallback
            // Create a new LoadOptions object and set its WarningCallback property. 
            LoadOptions loadOptions = new LoadOptions { WarningCallback = new DocumentLoadingWarningCallback() };
 
            Document doc = new Document(dataDir + "document.docx", loadOptions);
            //ExEnd:LoadOptionsWarningCallback
        }

        //ExStart:DocumentLoadingWarningCallback
        public class DocumentLoadingWarningCallback : IWarningCallback
        {
            public void Warning(WarningInfo info)
            {
                // Prints warnings and their details as they arise during document loading.
                Console.WriteLine($"WARNING: {info.WarningType}, source: {info.Source}");
                Console.WriteLine($"\tDescription: {info.Description}");
            }
        }
        //ExEnd:DocumentLoadingWarningCallback
        
        public static void LoadOptionsResourceLoadingCallback(string dataDir)
        {
            //ExStart:LoadOptionsResourceLoadingCallback
            // Create a new LoadOptions object and set its ResourceLoadingCallback attribute as an instance of our IResourceLoadingCallback implementation
            LoadOptions loadOptions = new LoadOptions { ResourceLoadingCallback = new HtmlLinkedResourceLoadingCallback() };
 
            // When we open an Html document, external resources such as references to CSS stylesheet files and external images
            // will be handled in a custom manner by the loading callback as the document is loaded
            Document doc = new Document(dataDir + "Images.html", loadOptions);
            doc.Save(dataDir + "Document.LoadOptionsCallback_out.pdf");
            //ExEnd:LoadOptionsResourceLoadingCallback
        }

        //ExStart:HtmlLinkedResourceLoadingCallback
        private class HtmlLinkedResourceLoadingCallback : IResourceLoadingCallback
        {
            public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
            {
                switch (args.ResourceType)
                {
                    case ResourceType.CssStyleSheet:
                    {
                        Console.WriteLine($"External CSS Stylesheet found upon loading: {args.OriginalUri}");
 
                        // CSS file will don't used in the document
                        return ResourceLoadingAction.Skip;
                    }
                    case ResourceType.Image:
                    {
                        // Replaces all images with a substitute
                        const string newImageFilename = "Logo.jpg";
                        Console.WriteLine($"\tImage will be substituted with: {newImageFilename}");
                        Image newImage = Image.FromFile(RunExamples.GetDataDir_QuickStart() + newImageFilename);
                        ImageConverter converter = new ImageConverter();
                        byte[] imageBytes = (byte[])converter.ConvertTo(newImage, typeof(byte[]));
                        args.SetData(imageBytes);
 
                        // New images will be used instead of presented in the document
                        return ResourceLoadingAction.UserProvided;
                    }
                    case ResourceType.Document:
                    {
                        Console.WriteLine($"External document found upon loading: {args.OriginalUri}");
 
                        // Will be used as usual
                        return ResourceLoadingAction.Default;
                    }
                    default:
                        throw new InvalidOperationException("Unexpected ResourceType value.");
                }
            }
        }
        //ExEnd:HtmlLinkedResourceLoadingCallback

        public static void LoadOptionsEncoding(string dataDir)
        {
            //ExStart:LoadOptionsEncoding
            // Set the Encoding attribute in a LoadOptions object to override the automatically chosen encoding with the one we know to be correct
            LoadOptions loadOptions = new LoadOptions { Encoding = Encoding.UTF7 };
            Document doc = new Document(dataDir + "Encoded in UTF-7.txt", loadOptions);
            //ExEnd:LoadOptionsEncoding
        }

        public static void SkipPdfImages(string dataDir)
        {
            //ExStart:SkipPdfImages
            PdfLoadOptions loadOptions = new PdfLoadOptions();
            loadOptions.SkipPdfImages = true;

            Document doc = new Document(dataDir + "in.pdf", loadOptions);
            //ExEnd:SkipPdfImages
        }

        public static void ConvertMetafilesToPng(string dataDir)
        {
            //ExStart:ConvertMetafilesToPng
            LoadOptions lo = new LoadOptions();
            lo.ConvertMetafilesToPng = true;

            Document doc = new Document(dataDir + "in.doc", lo);
            //ExEnd:ConvertMetafilesToPng
        }
    }
}
