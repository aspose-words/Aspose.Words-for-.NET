using System;
using System.Drawing;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Load_Options
{
    public class WorkingWithLoadOptions : DocsExamplesBase
    {
        [Test]
        public void UpdateDirtyFields()
        {
            //ExStart:UpdateDirtyFields
            LoadOptions loadOptions = new LoadOptions { UpdateDirtyFields = true };

            Document doc = new Document(MyDir + "Dirty field.docx", loadOptions);

            doc.Save(ArtifactsDir + "WorkingWithLoadOptions.UpdateDirtyFields.docx");
            //ExEnd:UpdateDirtyFields
        }

        [Test]
        public void LoadEncryptedDocument()
        {
            //ExStart:LoadSaveEncryptedDoc
            //ExStart:OpenEncryptedDocument
            Document doc = new Document(MyDir + "Encrypted.docx", new LoadOptions("docPassword"));
            //ExEnd:OpenEncryptedDocument

            doc.Save(ArtifactsDir + "WorkingWithLoadOptions.LoadAndSaveEncryptedOdt.odt", new OdtSaveOptions("newPassword"));
            //ExEnd:LoadSaveEncryptedDoc
        }

        [Test]
        public void ConvertShapeToOfficeMath()
        {
            //ExStart:ConvertShapeToOfficeMath
            LoadOptions loadOptions = new LoadOptions { ConvertShapeToOfficeMath = true };

            Document doc = new Document(MyDir + "Office math.docx", loadOptions);

            doc.Save(ArtifactsDir + "WorkingWithLoadOptions.ConvertShapeToOfficeMath.docx", SaveFormat.Docx);
            //ExEnd:ConvertShapeToOfficeMath
        }

        [Test]
        public void SetMsWordVersion()
        {
            //ExStart:SetMSWordVersion
            // Create a new LoadOptions object, which will load documents according to MS Word 2019 specification by default
            // and change the loading version to Microsoft Word 2010.
            LoadOptions loadOptions = new LoadOptions { MswVersion = MsWordVersion.Word2010 };
            
            Document doc = new Document(MyDir + "Document.docx", loadOptions);

            doc.Save(ArtifactsDir + "WorkingWithLoadOptions.SetMsWordVersion.docx");
            //ExEnd:SetMSWordVersion
        }

        [Test]
        public void UseTempFolder()
        {
            //ExStart:UseTempFolder  
            LoadOptions loadOptions = new LoadOptions { TempFolder = ArtifactsDir };

            Document doc = new Document(MyDir + "Document.docx", loadOptions);
            //ExEnd:UseTempFolder  
        }
        
        [Test]
        public void WarningCallback()
        {
            //ExStart:WarningCallback
            LoadOptions loadOptions = new LoadOptions { WarningCallback = new DocumentLoadingWarningCallback() };
            
            Document doc = new Document(MyDir + "Document.docx", loadOptions);
            //ExEnd:WarningCallback
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
        
#if NET462
        [Test]
        public void ResourceLoadingCallback()
        {
            //ExStart:ResourceLoadingCallback
            LoadOptions loadOptions = new LoadOptions { ResourceLoadingCallback = new HtmlLinkedResourceLoadingCallback() };

            // When we open an Html document, external resources such as references to CSS stylesheet files
            // and external images will be handled customarily by the loading callback as the document is loaded.
            Document doc = new Document(MyDir + "Images.html", loadOptions);

            doc.Save(ArtifactsDir + "WorkingWithLoadOptions.ResourceLoadingCallback.pdf");
            //ExEnd:ResourceLoadingCallback
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
 
                        // CSS file will don't used in the document.
                        return ResourceLoadingAction.Skip;
                    }
                    case ResourceType.Image:
                    {
                        // Replaces all images with a substitute.
                        Image newImage = Image.FromFile(ImagesDir + "Logo.jpg");
                        
                        ImageConverter converter = new ImageConverter();
                        byte[] imageBytes = (byte[])converter.ConvertTo(newImage, typeof(byte[]));

                        args.SetData(imageBytes);
 
                        // New images will be used instead of presented in the document.
                        return ResourceLoadingAction.UserProvided;
                    }
                    case ResourceType.Document:
                    {
                        Console.WriteLine($"External document found upon loading: {args.OriginalUri}");
 
                        // Will be used as usual.
                        return ResourceLoadingAction.Default;
                    }
                    default:
                        throw new InvalidOperationException("Unexpected ResourceType value.");
                }
            }
        }
        //ExEnd:HtmlLinkedResourceLoadingCallback
#endif

        [Test]
        public void LoadWithEncoding()
        {
            //ExStart:LoadWithEncoding
            LoadOptions loadOptions = new LoadOptions { Encoding = Encoding.UTF7 };

            Document doc = new Document(MyDir + "Encoded in UTF-7.txt", loadOptions);
            //ExEnd:LoadWithEncoding
        }

        [Test]
        public void SkipPdfImages()
        {
            //ExStart:SkipPdfImages
            PdfLoadOptions loadOptions = new PdfLoadOptions { SkipPdfImages = true };

            Document doc = new Document(MyDir + "Pdf Document.pdf", loadOptions);
            //ExEnd:SkipPdfImages
        }

        [Test]
        public void ConvertMetafilesToPng()
        {
            //ExStart:ConvertMetafilesToPng
            LoadOptions loadOptions = new LoadOptions { ConvertMetafilesToPng = true };

            Document doc = new Document(MyDir + "WMF with image.docx", loadOptions);
            //ExEnd:ConvertMetafilesToPng
        }

        [Test]
        public void LoadChm()
        {
            //ExStart:LoadCHM
            LoadOptions loadOptions = new LoadOptions { Encoding = Encoding.GetEncoding("windows-1251") };

            Document doc = new Document(MyDir + "HTML help.chm", loadOptions);
            //ExEnd:LoadCHM
        }
    }
}