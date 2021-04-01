// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fonts;
using Aspose.Words.Loading;
using Aspose.Words.Settings;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    class ExLoadOptions : ApiExampleBase
    {
#if NET462 || MAC || JAVA
        //ExStart
        //ExFor:LoadOptions.ResourceLoadingCallback
        //ExSummary:Shows how to handle external resources when loading Html documents.
        [Test] //ExSkip
        public void LoadOptionsCallback()
        {
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ResourceLoadingCallback = new HtmlLinkedResourceLoadingCallback();

            // When we load the document, our callback will handle linked resources such as CSS stylesheets and images.
            Document doc = new Document(MyDir + "Images.html", loadOptions);
            doc.Save(ArtifactsDir + "LoadOptions.LoadOptionsCallback.pdf");
        }

        /// <summary>
        /// Prints the filenames of all external stylesheets and substitutes all images of a loaded html document.
        /// </summary>
        private class HtmlLinkedResourceLoadingCallback : IResourceLoadingCallback
        {
            public ResourceLoadingAction ResourceLoading(ResourceLoadingArgs args)
            {
                switch (args.ResourceType)
                {
                    case ResourceType.CssStyleSheet:
                        Console.WriteLine($"External CSS Stylesheet found upon loading: {args.OriginalUri}");
                        return ResourceLoadingAction.Default;
                    case ResourceType.Image:
                        Console.WriteLine($"External Image found upon loading: {args.OriginalUri}");

                        const string newImageFilename = "Logo.jpg";
                        Console.WriteLine($"\tImage will be substituted with: {newImageFilename}");

                        Image newImage = Image.FromFile(ImageDir + newImageFilename);

                        ImageConverter converter = new ImageConverter();
                        byte[] imageBytes = (byte[])converter.ConvertTo(newImage, typeof(byte[]));
                        args.SetData(imageBytes);

                        return ResourceLoadingAction.UserProvided;
                }

                return ResourceLoadingAction.Default;
            }
        }
        //ExEnd
#endif

        [TestCase(true)]
        [TestCase(false)]
        public void ConvertShapeToOfficeMath(bool isConvertShapeToOfficeMath)
        {
            //ExStart
            //ExFor:LoadOptions.ConvertShapeToOfficeMath
            //ExSummary:Shows how to convert EquationXML shapes to Office Math objects.
            LoadOptions loadOptions = new LoadOptions();

            // Use this flag to specify whether to convert the shapes with EquationXML attributes
            // to Office Math objects and then load the document.
            loadOptions.ConvertShapeToOfficeMath = isConvertShapeToOfficeMath;

            Document doc = new Document(MyDir + "Math shapes.docx", loadOptions);

            if (isConvertShapeToOfficeMath)
            {
                Assert.AreEqual(16, doc.GetChildNodes(NodeType.Shape, true).Count);
                Assert.AreEqual(34, doc.GetChildNodes(NodeType.OfficeMath, true).Count);
            }
            else
            {
                Assert.AreEqual(24, doc.GetChildNodes(NodeType.Shape, true).Count);
                Assert.AreEqual(0, doc.GetChildNodes(NodeType.OfficeMath, true).Count);
            }
            //ExEnd
        }

        [Test]
        public void SetEncoding()
        {
            //ExStart
            //ExFor:LoadOptions.Encoding
            //ExSummary:Shows how to set the encoding with which to open a document.
            // A FileFormatInfo object will detect this file as being encoded in something other than UTF-7.
            FileFormatInfo fileFormatInfo = FileFormatUtil.DetectFileFormat(MyDir + "Encoded in UTF-7.txt");

            Assert.AreNotEqual(Encoding.UTF7, fileFormatInfo.Encoding);

            // If we load the document with no loading configurations, Aspose.Words will detect its encoding as UTF-8.
            Document doc = new Document(MyDir + "Encoded in UTF-7.txt");

            // The contents, parsed in UTF-8, create a valid string.
            // However, knowing that the file is in UTF-7, we can see that the result is incorrect.
            Assert.AreEqual("Hello world+ACE-", doc.ToString(SaveFormat.Text).Trim());

            // In cases of ambiguous encoding such as this one, we can set a specific encoding variant
            // to parse the file within a LoadOptions object.
            LoadOptions loadOptions = new LoadOptions
            {
                Encoding = Encoding.UTF7
            };

            // Load the document while passing the LoadOptions object, then verify the document's contents.
            doc = new Document(MyDir + "Encoded in UTF-7.txt", loadOptions);

            Assert.AreEqual("Hello world!", doc.ToString(SaveFormat.Text).Trim());
            //ExEnd
        }

        [Test]
        public void FontSettings()
        {
            //ExStart
            //ExFor:LoadOptions.FontSettings
            //ExSummary:Shows how to apply font substitution settings while loading a document. 
            // Create a FontSettings object that will substitute the "Times New Roman" font
            // with the font "Arvo" from our "MyFonts" folder.
            FontSettings fontSettings = new FontSettings();
            fontSettings.SetFontsFolder(FontsDir, false);
            fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Times New Roman", "Arvo");

            // Set that FontSettings object as a property of a newly created LoadOptions object.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.FontSettings = fontSettings;

            // Load the document, then render it as a PDF with the font substitution.
            Document doc = new Document(MyDir + "Document.docx", loadOptions);

            doc.Save(ArtifactsDir + "LoadOptions.FontSettings.pdf");
            //ExEnd
        }

        [Test]
        public void LoadOptionsMswVersion()
        {
            //ExStart
            //ExFor:LoadOptions.MswVersion
            //ExSummary:Shows how to emulate the loading procedure of a specific Microsoft Word version during document loading.
            // By default, Aspose.Words load documents according to Microsoft Word 2019 specification.
            LoadOptions loadOptions = new LoadOptions();
            
            Assert.AreEqual(MsWordVersion.Word2019, loadOptions.MswVersion);

            // This document is missing the default paragraph formatting style.
            // This default style will be regenerated when we load the document either with Microsoft Word or Aspose.Words.
            loadOptions.MswVersion = MsWordVersion.Word2007;
            Document doc = new Document(MyDir + "Document.docx", loadOptions);

            // The style's line spacing will have this value when loaded by Microsoft Word 2007 specification.
            Assert.AreEqual(12.95d, doc.Styles.DefaultParagraphFormat.LineSpacing, 0.01d);
            //ExEnd
        }

        //ExStart
        //ExFor:LoadOptions.WarningCallback
        //ExSummary:Shows how to print and store warnings that occur during document loading.
        [Test] //ExSkip
        public void LoadOptionsWarningCallback()
        {
            // Create a new LoadOptions object and set its WarningCallback attribute
            // as an instance of our IWarningCallback implementation.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.WarningCallback = new DocumentLoadingWarningCallback();

            // Our callback will print all warnings that come up during the load operation.
            Document doc = new Document(MyDir + "Document.docx", loadOptions);

            List<WarningInfo> warnings = ((DocumentLoadingWarningCallback)loadOptions.WarningCallback).GetWarnings();
            Assert.AreEqual(3, warnings.Count);
            TestLoadOptionsWarningCallback(warnings); //ExSkip
        }

        /// <summary>
        /// IWarningCallback that prints warnings and their details as they arise during document loading.
        /// </summary>
        private class DocumentLoadingWarningCallback : IWarningCallback
        {
            public void Warning(WarningInfo info)
            {
                Console.WriteLine($"Warning: {info.WarningType}");
                Console.WriteLine($"\tSource: {info.Source}");
                Console.WriteLine($"\tDescription: {info.Description}");
                mWarnings.Add(info);
            }

            public List<WarningInfo> GetWarnings()
            {
                return mWarnings;
            }

            private readonly List<WarningInfo> mWarnings = new List<WarningInfo>();
        }
        //ExEnd

        private static void TestLoadOptionsWarningCallback(List<WarningInfo> warnings)
        {
            Assert.AreEqual(WarningType.UnexpectedContent, warnings[0].WarningType);
            Assert.AreEqual(WarningSource.Docx, warnings[0].Source);
            Assert.AreEqual("3F01", warnings[0].Description);

            Assert.AreEqual(WarningType.MinorFormattingLoss, warnings[1].WarningType);
            Assert.AreEqual(WarningSource.Docx, warnings[1].Source);
            Assert.AreEqual("Import of element 'shapedefaults' is not supported in Docx format by Aspose.Words.", warnings[1].Description);

            Assert.AreEqual(WarningType.MinorFormattingLoss, warnings[2].WarningType);
            Assert.AreEqual(WarningSource.Docx, warnings[2].Source);
            Assert.AreEqual("Import of element 'extraClrSchemeLst' is not supported in Docx format by Aspose.Words.", warnings[2].Description);
        }

        [Test]
        public void TempFolder()
        {
            //ExStart
            //ExFor:LoadOptions.TempFolder
            //ExSummary:Shows how to use the hard drive instead of memory when loading a document.
            // When we load a document, various elements are temporarily stored in memory as the save operation occurs.
            // We can use this option to use a temporary folder in the local file system instead,
            // which will reduce our application's memory overhead.
            LoadOptions options = new LoadOptions();
            options.TempFolder = ArtifactsDir + "TempFiles";

            // The specified temporary folder must exist in the local file system before the load operation.
            Directory.CreateDirectory(options.TempFolder);

            Document doc = new Document(MyDir + "Document.docx", options);

            // The folder will persist with no residual contents from the load operation.
            Assert.That(Directory.GetFiles(options.TempFolder), Is.Empty);
            //ExEnd
        }

        [Test]
        public void AddEditingLanguage()
        {
            //ExStart
            //ExFor:LanguagePreferences
            //ExFor:LanguagePreferences.AddEditingLanguage(EditingLanguage)
            //ExFor:LoadOptions.LanguagePreferences
            //ExFor:EditingLanguage
            //ExSummary:Shows how to apply language preferences when loading a document.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LanguagePreferences.AddEditingLanguage(EditingLanguage.Japanese);

            Document doc = new Document(MyDir + "No default editing language.docx", loadOptions);

            int localeIdFarEast = doc.Styles.DefaultFont.LocaleIdFarEast;
            Console.WriteLine(localeIdFarEast == (int)EditingLanguage.Japanese
                ? "The document either has no any FarEast language set in defaults or it was set to Japanese originally."
                : "The document default FarEast language was set to another than Japanese language originally, so it is not overridden.");
            //ExEnd

            Assert.AreEqual((int)EditingLanguage.Japanese, doc.Styles.DefaultFont.LocaleIdFarEast);

            doc = new Document(MyDir + "No default editing language.docx");

            Assert.AreEqual((int)EditingLanguage.EnglishUS, doc.Styles.DefaultFont.LocaleIdFarEast);
        }

        [Test]
        public void SetEditingLanguageAsDefault()
        {
            //ExStart
            //ExFor:LanguagePreferences.DefaultEditingLanguage
            //ExSummary:Shows how set a default language when loading a document.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.LanguagePreferences.DefaultEditingLanguage = EditingLanguage.Russian;

            Document doc = new Document(MyDir + "No default editing language.docx", loadOptions);

            int localeId = doc.Styles.DefaultFont.LocaleId;
            Console.WriteLine(localeId == (int)EditingLanguage.Russian
                ? "The document either has no any language set in defaults or it was set to Russian originally."
                : "The document default language was set to another than Russian language originally, so it is not overridden.");
            //ExEnd

            Assert.AreEqual((int)EditingLanguage.Russian, doc.Styles.DefaultFont.LocaleId);

            doc = new Document(MyDir + "No default editing language.docx");

            Assert.AreEqual((int)EditingLanguage.EnglishUS, doc.Styles.DefaultFont.LocaleId);
        }

        [Test]
        public void ConvertMetafilesToPng()
        {
            //ExStart
            //ExFor:LoadOptions.ConvertMetafilesToPng
            //ExSummary:Shows how to convert WMF/EMF to PNG during loading document.
            Document doc = new Document();
    
            Shape shape = new Shape(doc, ShapeType.Image);
            shape.ImageData.SetImage(ImageDir + "Windows MetaFile.wmf");
            shape.Width = 100;
            shape.Height = 100;

            doc.FirstSection.Body.FirstParagraph.AppendChild(shape);

            doc.Save(ArtifactsDir + "Image.CreateImageDirectly.docx");

            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            TestUtil.VerifyImageInShape(1600, 1600, ImageType.Wmf, shape);

            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ConvertMetafilesToPng = true;

            doc = new Document(ArtifactsDir + "Image.CreateImageDirectly.docx", loadOptions);
            shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);

#if NET462
            TestUtil.VerifyImageInShape(1666, 1666, ImageType.Png, shape);
#elif NETCOREAPP2_1
            TestUtil.VerifyImageInShape(1666, 1666, ImageType.Png, shape);
#endif
            //ExEnd
        }

        [Test]
        public void OpenChmFile()
        {
            FileFormatInfo info = FileFormatUtil.DetectFileFormat(MyDir + "HTML help.chm");
            Assert.AreEqual(info.LoadFormat, LoadFormat.Chm);

            LoadOptions loadOptions = new LoadOptions();
            loadOptions.Encoding = Encoding.GetEncoding("windows-1251");

            Document doc = new Document(MyDir + "HTML help.chm", loadOptions);
        }
    }
}
