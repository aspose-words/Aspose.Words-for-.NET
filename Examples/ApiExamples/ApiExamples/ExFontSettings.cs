// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Loading;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExFontSettings : ApiExampleBase
    {
        [Test]
        public void DefaultFontInstance()
        {
            //ExStart
            //ExFor:FontSettings.DefaultInstance
            //ExSummary:Shows how to configure the default font settings instance.
            // Configure the default font settings instance to use the "Courier New" font
            // as a backup substitute when we attempt to use an unknown font.
            FontSettings.DefaultInstance.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Courier New";

            Assert.That(FontSettings.DefaultInstance.SubstitutionSettings.DefaultFontSubstitution.Enabled, Is.True);

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Non-existent font";
            builder.Write("Hello world!");

            // This document does not have a FontSettings configuration. When we render the document,
            // the default FontSettings instance will resolve the missing font.
            // Aspose.Words will use "Courier New" to render text that uses the unknown font.
            Assert.That(doc.FontSettings, Is.Null);

            doc.Save(ArtifactsDir + "FontSettings.DefaultFontInstance.pdf");
            //ExEnd
        }

        [Test]
        public void DefaultFontName()
        {
            //ExStart
            //ExFor:DefaultFontSubstitutionRule.DefaultFontName
            //ExSummary:Shows how to specify a default font.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Arial";
            builder.Writeln("Hello world!");
            builder.Font.Name = "Arvo";
            builder.Writeln("The quick brown fox jumps over the lazy dog.");

            FontSourceBase[] fontSources = FontSettings.DefaultInstance.GetFontsSources();

            // The font sources that the document uses contain the font "Arial", but not "Arvo".
            Assert.That(fontSources.Length, Is.EqualTo(1));
            Assert.That(fontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arial"), Is.True);
            Assert.That(fontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arvo"), Is.False);

            // Set the "DefaultFontName" property to "Courier New" to,
            // while rendering the document, apply that font in all cases when another font is not available. 
            FontSettings.DefaultInstance.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Courier New";

            Assert.That(fontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Courier New"), Is.True);

            // Aspose.Words will now use the default font in place of any missing fonts during any rendering calls.
            doc.Save(ArtifactsDir + "FontSettings.DefaultFontName.pdf");
            //ExEnd
        }

        [Test]
        public void UpdatePageLayoutWarnings()
        {
            // Store the font sources currently used so we can restore them later
            FontSourceBase[] originalFontSources = FontSettings.DefaultInstance.GetFontsSources();

            // Load the document to render
            Document doc = new Document(MyDir + "Document.docx");

            // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class
            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.WarningCallback = callback;

            // We can choose the default font to use in the case of any missing fonts
            FontSettings.DefaultInstance.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";

            // For testing we will set Aspose.Words to look for fonts only in a folder which does not exist. Since Aspose.Words won't
            // find any fonts in the specified directory, then during rendering the fonts in the document will be substituted with the default 
            // font specified under FontSettings.DefaultFontName. We can pick up on this substitution using our callback
            FontSettings.DefaultInstance.SetFontsFolder(string.Empty, false);

            // When you call UpdatePageLayout the document is rendered in memory. Any warnings that occurred during rendering
            // are stored until the document save and then sent to the appropriate WarningCallback
            doc.UpdatePageLayout();

            // Even though the document was rendered previously, any save warnings are notified to the user during document save
            doc.Save(ArtifactsDir + "FontSettings.UpdatePageLayoutWarnings.pdf");

            Assert.That(callback.FontWarnings.Count > 0, Is.True);
            Assert.That(callback.FontWarnings[0].WarningType == WarningType.FontSubstitution, Is.True);
            Assert.That(callback.FontWarnings[0].Description.Contains("has not been found"), Is.True);

            // Restore default fonts
            FontSettings.DefaultInstance.SetFontsSources(originalFontSources);
        }

        public class HandleDocumentWarnings : IWarningCallback
        {
            /// <summary>
            /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
            /// potential issue during document processing. The callback can be set to listen for warnings generated during document
            /// load and/or document save.
            /// </summary>
            public void Warning(WarningInfo info)
            {
                // We are only interested in fonts being substituted
                if (info.WarningType == WarningType.FontSubstitution)
                {
                    Console.WriteLine("Font substitution: " + info.Description);
                    FontWarnings.Warning(info);
                }
            }

            public WarningInfoCollection FontWarnings = new WarningInfoCollection();
        }

        //ExStart
        //ExFor:IWarningCallback
        //ExFor:DocumentBase.WarningCallback
        //ExFor:FontSettings.DefaultInstance
        //ExSummary:Shows how to use the IWarningCallback interface to monitor font substitution warnings.
        [Test] //ExSkip
        public void SubstitutionWarning()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Times New Roman";
            builder.Writeln("Hello world!");

            FontSubstitutionWarningCollector callback = new FontSubstitutionWarningCollector();
            doc.WarningCallback = callback;

            // Store the current collection of font sources, which will be the default font source for every document
            // for which we do not specify a different font source.
            FontSourceBase[] originalFontSources = FontSettings.DefaultInstance.GetFontsSources();

            // For testing purposes, we will set Aspose.Words to look for fonts only in a folder that does not exist.
            FontSettings.DefaultInstance.SetFontsFolder(string.Empty, false);

            // When rendering the document, there will be no place to find the "Times New Roman" font.
            // This will cause a font substitution warning, which our callback will detect.
            doc.Save(ArtifactsDir + "FontSettings.SubstitutionWarning.pdf");

            FontSettings.DefaultInstance.SetFontsSources(originalFontSources);

            Assert.That(callback.FontSubstitutionWarnings.Count, Is.EqualTo(1)); //ExSkip
            Assert.That(callback.FontSubstitutionWarnings[0].WarningType == WarningType.FontSubstitution, Is.True);
            Assert.That(callback.FontSubstitutionWarnings[0].Description
                .Equals(
                    "Font 'Times New Roman' has not been found. Using 'Fanwood' font instead. Reason: first available font."), Is.True);
        }

        private class FontSubstitutionWarningCollector : IWarningCallback
        {
            /// <summary>
            /// Called every time a warning occurs during loading/saving.
            /// </summary>
            public void Warning(WarningInfo info)
            {
                if (info.WarningType == WarningType.FontSubstitution)
                    FontSubstitutionWarnings.Warning(info);
            }

            public WarningInfoCollection FontSubstitutionWarnings = new WarningInfoCollection();
        }
        //ExEnd

        //ExStart
        //ExFor:FontSourceBase.WarningCallback
        //ExSummary:Shows how to call warning callback when the font sources working with.
        [Test]//ExSkip
        public void FontSourceWarning()
        {
            FontSettings settings = new FontSettings();
            settings.SetFontsFolder("bad folder?", false);

            FontSourceBase source = settings.GetFontsSources()[0];
            FontSourceWarningCollector callback = new FontSourceWarningCollector();
            source.WarningCallback = callback;

            // Get the list of fonts to call warning callback.
            IList<PhysicalFontInfo> fontInfos = source.GetAvailableFonts();

            Assert.That(callback.FontSubstitutionWarnings[0].Description
                .Contains("Error loading font from the folder \"bad folder?\""), Is.True);
        }

        private class FontSourceWarningCollector : IWarningCallback
        {
            /// <summary>
            /// Called every time a warning occurs during processing of font source.
            /// </summary>
            public void Warning(WarningInfo info)
            {
                FontSubstitutionWarnings.Warning(info);
            }

            public readonly WarningInfoCollection FontSubstitutionWarnings = new WarningInfoCollection();
        }
        //ExEnd

        
        [Test]
        public void EnableFontSubstitution()
        {
            //ExStart
            //ExFor:FontInfoSubstitutionRule
            //ExFor:FontSubstitutionSettings.FontInfoSubstitution
            //ExFor:LayoutOptions.KeepOriginalFontMetrics
            //ExFor:IWarningCallback
            //ExFor:IWarningCallback.Warning(WarningInfo)
            //ExFor:WarningInfo
            //ExFor:WarningInfo.Description
            //ExFor:WarningInfo.WarningType
            //ExFor:WarningInfoCollection
            //ExFor:WarningInfoCollection.Warning(WarningInfo)
            //ExFor:WarningInfoCollection.Clear
            //ExFor:WarningType
            //ExFor:DocumentBase.WarningCallback
            //ExSummary:Shows how to set the property for finding the closest match for a missing font from the available font sources.
            // Open a document that contains text formatted with a font that does not exist in any of our font sources.
            Document doc = new Document(MyDir + "Missing font.docx");

            // Assign a callback for handling font substitution warnings.
            WarningInfoCollection warningCollector = new WarningInfoCollection();
            doc.WarningCallback = warningCollector;

            // Set a default font name and enable font substitution.
            FontSettings fontSettings = new FontSettings();
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
            fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = true;

            // Original font metrics should be used after font substitution.
            doc.LayoutOptions.KeepOriginalFontMetrics = true;

            // We will get a font substitution warning if we save a document with a missing font.
            doc.FontSettings = fontSettings;
            doc.Save(ArtifactsDir + "FontSettings.EnableFontSubstitution.pdf");

            foreach (WarningInfo info in warningCollector)
            {
                if (info.WarningType == WarningType.FontSubstitution)
                    Console.WriteLine(info.Description);
            }
            //ExEnd

            // We can also verify warnings in the collection and clear them.
            Assert.That(warningCollector[0].Source, Is.EqualTo(WarningSource.Layout));
            Assert.That(warningCollector[0].Description, Is.EqualTo("Font '28 Days Later' has not been found. Using 'Calibri' font instead. Reason: alternative name from document."));

            warningCollector.Clear();

            Assert.That(warningCollector.Count, Is.EqualTo(0));
        }
        

        [Test]
        public void SubstitutionWarningsClosestMatch()
        {
            Document doc = new Document(MyDir + "Bullet points with alternative font.docx");

            WarningInfoCollection callback = new WarningInfoCollection();
            doc.WarningCallback = callback;

            doc.Save(ArtifactsDir + "FontSettings.SubstitutionWarningsClosestMatch.pdf");

            Assert.That(callback[0].Description
                .Equals(
                    "Font \'SymbolPS\' has not been found. Using \'Wingdings\' font instead. Reason: font info substitution."), Is.True);
        }

        [Test]
        public void DisableFontSubstitution()
        {
            Document doc = new Document(MyDir + "Missing font.docx");

            WarningInfoCollection callback = new WarningInfoCollection();
            doc.WarningCallback = callback;

            FontSettings fontSettings = new FontSettings();
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
            fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = false;

            doc.FontSettings = fontSettings;
            doc.Save(ArtifactsDir + "FontSettings.DisableFontSubstitution.pdf");

            Regex reg = new Regex(
                "Font '28 Days Later' has not been found. Using (.*) font instead. Reason: default font setting.");

            foreach (WarningInfo fontWarning in callback)
            {
                Match match = reg.Match(fontWarning.Description);
                if (match.Success)
                {
                    Assert.Pass();
                }
            }
        }

        [Test]
        [Category("SkipMono")]
        public void SubstitutionWarnings()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            WarningInfoCollection callback = new WarningInfoCollection();
            doc.WarningCallback = callback;

            FontSettings fontSettings = new FontSettings();
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
            fontSettings.SetFontsFolder(FontsDir, false);
            fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Arial", "Arvo", "Slab");

            doc.FontSettings = fontSettings;
            doc.Save(ArtifactsDir + "FontSettings.SubstitutionWarnings.pdf");

            Assert.That(callback[0].Description, Is.EqualTo("Font \'Arial\' has not been found. Using \'Arvo\' font instead. Reason: table substitution."));
            Assert.That(callback[1].Description, Is.EqualTo("Font \'Times New Roman\' has not been found. Using \'M+ 2m\' font instead. Reason: font info substitution."));
        }

        [Test]
        public void GetSubstitutionWithoutSuffixes()
        {
            Document doc = new Document(MyDir + "Get substitution without suffixes.docx");

            FontSourceBase[] originalFontSources = FontSettings.DefaultInstance.GetFontsSources();

            WarningInfoCollection substitutionWarningHandler = new WarningInfoCollection();
            doc.WarningCallback = substitutionWarningHandler;

            List<FontSourceBase> fontSources = new List<FontSourceBase>(FontSettings.DefaultInstance.GetFontsSources());
            FolderFontSource folderFontSource = new FolderFontSource(FontsDir, true);
            fontSources.Add(folderFontSource);

            FontSourceBase[] updatedFontSources = fontSources.ToArray();
            FontSettings.DefaultInstance.SetFontsSources(updatedFontSources);

            doc.Save(ArtifactsDir + "Font.GetSubstitutionWithoutSuffixes.pdf");

            Assert.That(substitutionWarningHandler[0].Description, Is.EqualTo("Font 'DINOT-Regular' has not been found. Using 'DINOT' font instead. Reason: font name substitution."));

            FontSettings.DefaultInstance.SetFontsSources(originalFontSources);
        }

        [Test]
        public void FontSourceFile()
        {
            //ExStart
            //ExFor:FileFontSource
            //ExFor:FileFontSource.#ctor(String)
            //ExFor:FileFontSource.#ctor(String, Int32)
            //ExFor:FileFontSource.FilePath
            //ExFor:FileFontSource.Type
            //ExFor:FontSourceBase
            //ExFor:FontSourceBase.Priority
            //ExFor:FontSourceBase.Type
            //ExFor:FontSourceType
            //ExSummary:Shows how to use a font file in the local file system as a font source.
            FileFontSource fileFontSource = new FileFontSource(MyDir + "Alte DIN 1451 Mittelschrift.ttf", 0);

            Document doc = new Document();
            doc.FontSettings = new FontSettings();
            doc.FontSettings.SetFontsSources(new FontSourceBase[] {fileFontSource});

            Assert.That(fileFontSource.FilePath, Is.EqualTo(MyDir + "Alte DIN 1451 Mittelschrift.ttf"));
            Assert.That(fileFontSource.Type, Is.EqualTo(FontSourceType.FontFile));
            Assert.That(fileFontSource.Priority, Is.EqualTo(0));
            //ExEnd
        }

        [Test]
        public void FontSourceFolder()
        {
            //ExStart
            //ExFor:FolderFontSource
            //ExFor:FolderFontSource.#ctor(String, Boolean)
            //ExFor:FolderFontSource.#ctor(String, Boolean, Int32)
            //ExFor:FolderFontSource.FolderPath
            //ExFor:FolderFontSource.ScanSubfolders
            //ExFor:FolderFontSource.Type
            //ExSummary:Shows how to use a local system folder which contains fonts as a font source.

            // Create a font source from a folder that contains font files.
            FolderFontSource folderFontSource = new FolderFontSource(FontsDir, false, 1);

            Document doc = new Document();
            doc.FontSettings = new FontSettings();
            doc.FontSettings.SetFontsSources(new FontSourceBase[] {folderFontSource});

            Assert.That(folderFontSource.FolderPath, Is.EqualTo(FontsDir));
            Assert.That(folderFontSource.ScanSubfolders, Is.EqualTo(false));
            Assert.That(folderFontSource.Type, Is.EqualTo(FontSourceType.FontsFolder));
            Assert.That(folderFontSource.Priority, Is.EqualTo(1));
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void SetFontsFolder(bool recursive)
        {
            //ExStart
            //ExFor:FontSettings
            //ExFor:FontSettings.SetFontsFolder(String, Boolean)
            //ExSummary:Shows how to set a font source directory.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Arvo";
            builder.Writeln("Hello world!");
            builder.Font.Name = "Amethysta";
            builder.Writeln("The quick brown fox jumps over the lazy dog.");

            // Our font sources do not contain the font that we have used for text in this document.
            // If we use these font settings while rendering this document,
            // Aspose.Words will apply a fallback font to text which has a font that Aspose.Words cannot locate.
            FontSourceBase[] originalFontSources = FontSettings.DefaultInstance.GetFontsSources();

            Assert.That(originalFontSources.Length, Is.EqualTo(1));
            Assert.That(originalFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arial"), Is.True);

            // The default font sources are missing the two fonts that we are using in this document.
            Assert.That(originalFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arvo"), Is.False);
            Assert.That(originalFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Amethysta"), Is.False);

            // Use the "SetFontsFolder" method to set a directory which will act as a new font source.
            // Pass "false" as the "recursive" argument to include fonts from all the font files that are in the directory
            // that we are passing in the first argument, but not include any fonts in any of that directory's subfolders.
            // Pass "true" as the "recursive" argument to include all font files in the directory that we are passing
            // in the first argument, as well as all the fonts in its subdirectories.
            FontSettings.DefaultInstance.SetFontsFolder(FontsDir, recursive);

            FontSourceBase[] newFontSources = FontSettings.DefaultInstance.GetFontsSources();

            Assert.That(newFontSources.Length, Is.EqualTo(1));
            Assert.That(newFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arial"), Is.False);
            Assert.That(newFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arvo"), Is.True);

            // The "Amethysta" font is in a subfolder of the font directory.
            if (recursive)
            {
                Assert.That(newFontSources[0].GetAvailableFonts().Count, Is.EqualTo(25));
                Assert.That(newFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Amethysta"), Is.True);
            }
            else
            {
                Assert.That(newFontSources[0].GetAvailableFonts().Count, Is.EqualTo(18));
                Assert.That(newFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Amethysta"), Is.False);
            }

            doc.Save(ArtifactsDir + "FontSettings.SetFontsFolder.pdf");

            // Restore the original font sources.
            FontSettings.DefaultInstance.SetFontsSources(originalFontSources);
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void SetFontsFolders(bool recursive)
        {
            //ExStart
            //ExFor:FontSettings
            //ExFor:FontSettings.SetFontsFolders(String[], Boolean)
            //ExSummary:Shows how to set multiple font source directories.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Amethysta";
            builder.Writeln("The quick brown fox jumps over the lazy dog.");
            builder.Font.Name = "Junction Light";
            builder.Writeln("The quick brown fox jumps over the lazy dog.");

            // Our font sources do not contain the font that we have used for text in this document.
            // If we use these font settings while rendering this document,
            // Aspose.Words will apply a fallback font to text which has a font that Aspose.Words cannot locate.
            FontSourceBase[] originalFontSources = FontSettings.DefaultInstance.GetFontsSources();

            Assert.That(originalFontSources.Length, Is.EqualTo(1));
            Assert.That(originalFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arial"), Is.True);

            // The default font sources are missing the two fonts that we are using in this document.
            Assert.That(originalFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Amethysta"), Is.False);
            Assert.That(originalFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Junction Light"), Is.False);

            // Use the "SetFontsFolders" method to create a font source from each font directory that we pass as the first argument.
            // Pass "false" as the "recursive" argument to include fonts from all the font files that are in the directories
            // that we are passing in the first argument, but not include any fonts from any of the directories' subfolders.
            // Pass "true" as the "recursive" argument to include all font files in the directories that we are passing
            // in the first argument, as well as all the fonts in their subdirectories.
            FontSettings.DefaultInstance.SetFontsFolders(new[] {FontsDir + "/Amethysta", FontsDir + "/Junction"},
                recursive);

            FontSourceBase[] newFontSources = FontSettings.DefaultInstance.GetFontsSources();

            Assert.That(newFontSources.Length, Is.EqualTo(2));
            Assert.That(newFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arial"), Is.False);
            Assert.That(newFontSources[0].GetAvailableFonts().Count, Is.EqualTo(1));
            Assert.That(newFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Amethysta"), Is.True);

            // The "Junction" folder itself contains no font files, but has subfolders that do.
            if (recursive)
            {
                Assert.That(newFontSources[1].GetAvailableFonts().Count, Is.EqualTo(6));
                Assert.That(newFontSources[1].GetAvailableFonts().Any(f => f.FullFontName == "Junction Light"), Is.True);
            }
            else
            {
                Assert.That(newFontSources[1].GetAvailableFonts().Count, Is.EqualTo(0));
            }

            doc.Save(ArtifactsDir + "FontSettings.SetFontsFolders.pdf");

            // Restore the original font sources.
            FontSettings.DefaultInstance.SetFontsSources(originalFontSources);
            //ExEnd
        }

        [Test]
        public void AddFontSource()
        {
            //ExStart
            //ExFor:FontSettings
            //ExFor:FontSettings.GetFontsSources()
            //ExFor:FontSettings.SetFontsSources(FontSourceBase[])
            //ExSummary:Shows how to add a font source to our existing font sources.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Arial";
            builder.Writeln("Hello world!");
            builder.Font.Name = "Amethysta";
            builder.Writeln("The quick brown fox jumps over the lazy dog.");
            builder.Font.Name = "Junction Light";
            builder.Writeln("The quick brown fox jumps over the lazy dog.");

            FontSourceBase[] originalFontSources = FontSettings.DefaultInstance.GetFontsSources();

            Assert.That(originalFontSources.Length, Is.EqualTo(1));

            Assert.That(originalFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arial"), Is.True);

            // The default font source is missing two of the fonts that we are using in our document.
            // When we save this document, Aspose.Words will apply fallback fonts to all text formatted with inaccessible fonts.
            Assert.That(originalFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Amethysta"), Is.False);
            Assert.That(originalFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Junction Light"), Is.False);

            // Create a font source from a folder that contains fonts.
            FolderFontSource folderFontSource = new FolderFontSource(FontsDir, true);

            // Apply a new array of font sources that contains the original font sources, as well as our custom fonts.
            FontSourceBase[] updatedFontSources = {originalFontSources[0], folderFontSource};
            FontSettings.DefaultInstance.SetFontsSources(updatedFontSources);

            // Verify that Aspose.Words has access to all required fonts before we render the document to PDF.
            updatedFontSources = FontSettings.DefaultInstance.GetFontsSources();

            Assert.That(updatedFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arial"), Is.True);
            Assert.That(updatedFontSources[1].GetAvailableFonts().Any(f => f.FullFontName == "Amethysta"), Is.True);
            Assert.That(updatedFontSources[1].GetAvailableFonts().Any(f => f.FullFontName == "Junction Light"), Is.True);

            doc.Save(ArtifactsDir + "FontSettings.AddFontSource.pdf");

            // Restore the original font sources.
            FontSettings.DefaultInstance.SetFontsSources(originalFontSources);
            //ExEnd
        }

        [Test]
        public void SetSpecifyFontFolder()
        {
            FontSettings fontSettings = new FontSettings();
            fontSettings.SetFontsFolder(FontsDir, false);

            // Using load options
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.FontSettings = fontSettings;

            Document doc = new Document(MyDir + "Rendering.docx", loadOptions);

            FolderFontSource folderSource = ((FolderFontSource) doc.FontSettings.GetFontsSources()[0]);

            Assert.That(folderSource.FolderPath, Is.EqualTo(FontsDir));
            Assert.That(folderSource.ScanSubfolders, Is.False);
        }

        [Test]
        public void TableSubstitution()
        {
            //ExStart
            //ExFor:Document.FontSettings
            //ExFor:TableSubstitutionRule.SetSubstitutes(String, String[])
            //ExSummary:Shows how set font substitution rules.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Arial";
            builder.Writeln("Hello world!");
            builder.Font.Name = "Amethysta";
            builder.Writeln("The quick brown fox jumps over the lazy dog.");

            FontSourceBase[] fontSources = FontSettings.DefaultInstance.GetFontsSources();

            // The default font sources contain the first font that the document uses.
            Assert.That(fontSources.Length, Is.EqualTo(1));
            Assert.That(fontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arial"), Is.True);

            // The second font, "Amethysta", is unavailable.
            Assert.That(fontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Amethysta"), Is.False);

            // We can configure a font substitution table which determines
            // which fonts Aspose.Words will use as substitutes for unavailable fonts.
            // Set two substitution fonts for "Amethysta": "Arvo", and "Courier New".
            // If the first substitute is unavailable, Aspose.Words attempts to use the second substitute, and so on.
            doc.FontSettings = new FontSettings();
            doc.FontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes(
                "Amethysta", new[] {"Arvo", "Courier New"});

            // "Amethysta" is unavailable, and the substitution rule states that the first font to use as a substitute is "Arvo". 
            Assert.That(fontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arvo"), Is.False);

            // "Arvo" is also unavailable, but "Courier New" is. 
            Assert.That(fontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Courier New"), Is.True);

            // The output document will display the text that uses the "Amethysta" font formatted with "Courier New".
            doc.Save(ArtifactsDir + "FontSettings.TableSubstitution.pdf");
            //ExEnd
        }

        [Test]
        public void SetSpecifyFontFolders()
        {
            FontSettings fontSettings = new FontSettings();
            fontSettings.SetFontsFolders(new string[] {FontsDir, @"C:\Windows\Fonts\"}, true);

            // Using load options
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.FontSettings = fontSettings;
            Document doc = new Document(MyDir + "Rendering.docx", loadOptions);

            FolderFontSource folderSource = ((FolderFontSource) doc.FontSettings.GetFontsSources()[0]);
            Assert.That(folderSource.FolderPath, Is.EqualTo(FontsDir));
            Assert.That(folderSource.ScanSubfolders, Is.True);

            folderSource = ((FolderFontSource) doc.FontSettings.GetFontsSources()[1]);
            Assert.That(folderSource.FolderPath, Is.EqualTo(@"C:\Windows\Fonts\"));
            Assert.That(folderSource.ScanSubfolders, Is.True);
        }

        [Test]
        public void AddFontSubstitutes()
        {
            FontSettings fontSettings = new FontSettings();
            fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes("Slab",
                new string[] {"Times New Roman", "Arial"});
            fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Arvo",
                new string[] {"Open Sans", "Arial"});

            Document doc = new Document(MyDir + "Rendering.docx");
            doc.FontSettings = fontSettings;

            string[] alternativeFonts = doc.FontSettings.SubstitutionSettings.TableSubstitution.GetSubstitutes("Slab")
                .ToArray();
            Assert.That(alternativeFonts, Is.EqualTo(new string[] {"Times New Roman", "Arial"}));

            alternativeFonts = doc.FontSettings.SubstitutionSettings.TableSubstitution.GetSubstitutes("Arvo").ToArray();
            Assert.That(alternativeFonts, Is.EqualTo(new string[] {"Open Sans", "Arial"}));
        }

        [Test]
        public void FontSourceMemory()
        {
            //ExStart
            //ExFor:MemoryFontSource
            //ExFor:MemoryFontSource.#ctor(Byte[])
            //ExFor:MemoryFontSource.#ctor(Byte[], Int32)
            //ExFor:MemoryFontSource.FontData
            //ExFor:MemoryFontSource.Type
            //ExSummary:Shows how to use a byte array with data from a font file as a font source.

            byte[] fontBytes = File.ReadAllBytes(MyDir + "Alte DIN 1451 Mittelschrift.ttf");
            MemoryFontSource memoryFontSource = new MemoryFontSource(fontBytes, 0);

            Document doc = new Document();
            doc.FontSettings = new FontSettings();
            doc.FontSettings.SetFontsSources(new FontSourceBase[] {memoryFontSource});

            Assert.That(memoryFontSource.Type, Is.EqualTo(FontSourceType.MemoryFont));
            Assert.That(memoryFontSource.Priority, Is.EqualTo(0));
            //ExEnd
        }

        [Test]
        public void FontSourceSystem()
        {
            //ExStart
            //ExFor:TableSubstitutionRule.AddSubstitutes(String, String[])
            //ExFor:FontSubstitutionRule.Enabled
            //ExFor:TableSubstitutionRule.GetSubstitutes(String)
            //ExFor:FontSettings.ResetFontSources
            //ExFor:FontSettings.SubstitutionSettings
            //ExFor:FontSubstitutionSettings
            //ExFor:FontSubstitutionSettings.FontNameSubstitution
            //ExFor:SystemFontSource
            //ExFor:SystemFontSource.#ctor
            //ExFor:SystemFontSource.#ctor(Int32)
            //ExFor:SystemFontSource.GetSystemFontFolders
            //ExFor:SystemFontSource.Type
            //ExSummary:Shows how to access a document's system font source and set font substitutes.
            Document doc = new Document();
            doc.FontSettings = new FontSettings();

            // By default, a blank document always contains a system font source.
            Assert.That(doc.FontSettings.GetFontsSources().Length, Is.EqualTo(1));

            SystemFontSource systemFontSource = (SystemFontSource) doc.FontSettings.GetFontsSources()[0];
            Assert.That(systemFontSource.Type, Is.EqualTo(FontSourceType.SystemFonts));
            Assert.That(systemFontSource.Priority, Is.EqualTo(0));

            PlatformID pid = Environment.OSVersion.Platform;
            bool isWindows = (pid == PlatformID.Win32NT) || (pid == PlatformID.Win32S) ||
                             (pid == PlatformID.Win32Windows) || (pid == PlatformID.WinCE);
            if (isWindows)
            {
                const string fontsPath = @"C:\WINDOWS\Fonts";
                Assert.That(SystemFontSource.GetSystemFontFolders().FirstOrDefault()?.ToLower(), Is.EqualTo(fontsPath.ToLower()));
            }

            foreach (string systemFontFolder in SystemFontSource.GetSystemFontFolders())
            {
                Console.WriteLine(systemFontFolder);
            }

            // Set a font that exists in the Windows Fonts directory as a substitute for one that does not.
            doc.FontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = true;
            doc.FontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Kreon-Regular", new[] {"Calibri"});

            Assert.That(doc.FontSettings.SubstitutionSettings.TableSubstitution.GetSubstitutes("Kreon-Regular").Count(), Is.EqualTo(1));
            Assert.That(doc.FontSettings.SubstitutionSettings.TableSubstitution.GetSubstitutes("Kreon-Regular").ToArray(), Does.Contain("Calibri"));

            // Alternatively, we could add a folder font source in which the corresponding folder contains the font.
            FolderFontSource folderFontSource = new FolderFontSource(FontsDir, false);
            doc.FontSettings.SetFontsSources(new FontSourceBase[] {systemFontSource, folderFontSource});
            Assert.That(doc.FontSettings.GetFontsSources().Length, Is.EqualTo(2));

            // Resetting the font sources still leaves us with the system font source as well as our substitutes.
            doc.FontSettings.ResetFontSources();

            Assert.That(doc.FontSettings.GetFontsSources().Length, Is.EqualTo(1));
            Assert.That(doc.FontSettings.GetFontsSources()[0].Type, Is.EqualTo(FontSourceType.SystemFonts));
            Assert.That(doc.FontSettings.SubstitutionSettings.TableSubstitution.GetSubstitutes("Kreon-Regular").Count(), Is.EqualTo(1));
            Assert.That(doc.FontSettings.SubstitutionSettings.FontNameSubstitution.Enabled, Is.True);
            //ExEnd
        }

        [Test]
        public void LoadFontFallbackSettingsFromFile()
        {
            //ExStart
            //ExFor:FontFallbackSettings.Load(String)
            //ExFor:FontFallbackSettings.Save(String)
            //ExSummary:Shows how to load and save font fallback settings to/from an XML document in the local file system.
            Document doc = new Document(MyDir + "Rendering.docx");

            // Load an XML document that defines a set of font fallback settings.
            FontSettings fontSettings = new FontSettings();
            fontSettings.FallbackSettings.Load(MyDir + "Font fallback rules.xml");

            doc.FontSettings = fontSettings;
            doc.Save(ArtifactsDir + "FontSettings.LoadFontFallbackSettingsFromFile.pdf");

            // Save our document's current font fallback settings as an XML document.
            doc.FontSettings.FallbackSettings.Save(ArtifactsDir + "FallbackSettings.xml");
            //ExEnd
        }

        [Test]
        public void LoadFontFallbackSettingsFromStream()
        {
            //ExStart
            //ExFor:FontFallbackSettings.Load(Stream)
            //ExFor:FontFallbackSettings.Save(Stream)
            //ExSummary:Shows how to load and save font fallback settings to/from a stream.
            Document doc = new Document(MyDir + "Rendering.docx");

            // Load an XML document that defines a set of font fallback settings.
            using (FileStream fontFallbackStream = new FileStream(MyDir + "Font fallback rules.xml", FileMode.Open))
            {
                FontSettings fontSettings = new FontSettings();
                fontSettings.FallbackSettings.Load(fontFallbackStream);

                doc.FontSettings = fontSettings;
            }

            doc.Save(ArtifactsDir + "FontSettings.LoadFontFallbackSettingsFromStream.pdf");

            // Use a stream to save our document's current font fallback settings as an XML document.
            using (FileStream fontFallbackStream =
                new FileStream(ArtifactsDir + "FallbackSettings.xml", FileMode.Create))
            {
                doc.FontSettings.FallbackSettings.Save(fontFallbackStream);
            }
            //ExEnd

            XmlDocument fallbackSettingsDoc = new XmlDocument();
            fallbackSettingsDoc.LoadXml(File.ReadAllText(ArtifactsDir + "FallbackSettings.xml"));
            XmlNamespaceManager manager = new XmlNamespaceManager(fallbackSettingsDoc.NameTable);
            manager.AddNamespace("aw", "Aspose.Words");

            XmlNodeList rules =
                fallbackSettingsDoc.SelectNodes("//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);

            Assert.That(rules[0].Attributes["Ranges"].Value, Is.EqualTo("0B80-0BFF"));
            Assert.That(rules[0].Attributes["FallbackFonts"].Value, Is.EqualTo("Vijaya"));

            Assert.That(rules[1].Attributes["Ranges"].Value, Is.EqualTo("1F300-1F64F"));
            Assert.That(rules[1].Attributes["FallbackFonts"].Value, Is.EqualTo("Segoe UI Emoji, Segoe UI Symbol"));

            Assert.That(rules[2].Attributes["Ranges"].Value, Is.EqualTo("2000-206F, 2070-209F, 20B9"));
            Assert.That(rules[2].Attributes["FallbackFonts"].Value, Is.EqualTo("Arial"));

            Assert.That(rules[3].Attributes["Ranges"].Value, Is.EqualTo("3040-309F"));
            Assert.That(rules[3].Attributes["FallbackFonts"].Value, Is.EqualTo("MS Gothic"));
            Assert.That(rules[3].Attributes["BaseFonts"].Value, Is.EqualTo("Times New Roman"));

            Assert.That(rules[4].Attributes["Ranges"].Value, Is.EqualTo("3040-309F"));
            Assert.That(rules[4].Attributes["FallbackFonts"].Value, Is.EqualTo("MS Mincho"));

            Assert.That(rules[5].Attributes["FallbackFonts"].Value, Is.EqualTo("Arial Unicode MS"));
        }

        [Test]
        public void LoadNotoFontsFallbackSettings()
        {
            //ExStart
            //ExFor:FontFallbackSettings.LoadNotoFallbackSettings
            //ExSummary:Shows how to add predefined font fallback settings for Google Noto fonts.
            FontSettings fontSettings = new FontSettings();

            // These are free fonts licensed under the SIL Open Font License.
            // We can download the fonts here:
            // https://www.google.com/get/noto/#sans-lgc
            fontSettings.SetFontsFolder(FontsDir + "Noto", false);

            // Note that the predefined settings only use Sans-style Noto fonts with regular weight. 
            // Some of the Noto fonts use advanced typography features.
            // Fonts featuring advanced typography may not be rendered correctly as Aspose.Words currently do not support them.
            fontSettings.FallbackSettings.LoadNotoFallbackSettings();
            fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = false;
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Noto Sans";

            Document doc = new Document();
            doc.FontSettings = fontSettings;
            //ExEnd
        }

        [Test]
        public void DefaultFontSubstitutionRule()
        {
            //ExStart
            //ExFor:DefaultFontSubstitutionRule
            //ExFor:DefaultFontSubstitutionRule.DefaultFontName
            //ExFor:FontSubstitutionSettings.DefaultFontSubstitution
            //ExSummary:Shows how to set the default font substitution rule.
            Document doc = new Document();
            FontSettings fontSettings = new FontSettings();
            doc.FontSettings = fontSettings;

            // Get the default substitution rule within FontSettings.
            // This rule will substitute all missing fonts with "Times New Roman".
            DefaultFontSubstitutionRule defaultFontSubstitutionRule =
                fontSettings.SubstitutionSettings.DefaultFontSubstitution;
            Assert.That(defaultFontSubstitutionRule.Enabled, Is.True);
            Assert.That(defaultFontSubstitutionRule.DefaultFontName, Is.EqualTo("Times New Roman"));

            // Set the default font substitute to "Courier New".
            defaultFontSubstitutionRule.DefaultFontName = "Courier New";

            // Using a document builder, add some text in a font that we do not have to see the substitution take place,
            // and then render the result in a PDF.
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Missing Font";
            builder.Writeln("Line written in a missing font, which will be substituted with Courier New.");

            doc.Save(ArtifactsDir + "FontSettings.DefaultFontSubstitutionRule.pdf");
            //ExEnd

            Assert.That(defaultFontSubstitutionRule.DefaultFontName, Is.EqualTo("Courier New"));
        }

        [Test]
        public void FontConfigSubstitution()
        {
            //ExStart
            //ExFor:FontConfigSubstitutionRule
            //ExFor:FontConfigSubstitutionRule.Enabled
            //ExFor:FontConfigSubstitutionRule.IsFontConfigAvailable
            //ExFor:FontConfigSubstitutionRule.ResetCache
            //ExFor:FontSubstitutionRule
            //ExFor:FontSubstitutionRule.Enabled
            //ExFor:FontSubstitutionSettings.FontConfigSubstitution
            //ExSummary:Shows operating system-dependent font config substitution.
            FontSettings fontSettings = new FontSettings();
            FontConfigSubstitutionRule fontConfigSubstitution =
                fontSettings.SubstitutionSettings.FontConfigSubstitution;

            bool isWindows = new[] {PlatformID.Win32NT, PlatformID.Win32S, PlatformID.Win32Windows, PlatformID.WinCE}
                .Any(p => Environment.OSVersion.Platform == p);

            // The FontConfigSubstitutionRule object works differently on Windows/non-Windows platforms.
            // On Windows, it is unavailable.
            if (isWindows)
            {
                Assert.That(fontConfigSubstitution.Enabled, Is.False);
                Assert.That(fontConfigSubstitution.IsFontConfigAvailable(), Is.False);
            }

            bool isLinuxOrMac =
                new[] {PlatformID.Unix, PlatformID.MacOSX}.Any(p => Environment.OSVersion.Platform == p);

            // On Linux/Mac, we will have access to it, and will be able to perform operations.
            if (isLinuxOrMac)
            {
                Assert.That(fontConfigSubstitution.Enabled, Is.True);
                Assert.That(fontConfigSubstitution.IsFontConfigAvailable(), Is.True);

                fontConfigSubstitution.ResetCache();
            }

            //ExEnd
        }

        [Test]
        public void FallbackSettings()
        {
            //ExStart
            //ExFor:FontFallbackSettings.LoadMsOfficeFallbackSettings
            //ExFor:FontFallbackSettings.LoadNotoFallbackSettings
            //ExSummary:Shows how to load pre-defined fallback font settings.
            Document doc = new Document();

            FontSettings fontSettings = new FontSettings();
            doc.FontSettings = fontSettings;
            FontFallbackSettings fontFallbackSettings = fontSettings.FallbackSettings;

            // Save the default fallback font scheme to an XML document.
            // For example, one of the elements has a value of "0C00-0C7F" for Range and a corresponding "Vani" value for FallbackFonts.
            // This means that if the font some text is using does not have symbols for the 0x0C00-0x0C7F Unicode block,
            // the fallback scheme will use symbols from the "Vani" font substitute.
            fontFallbackSettings.Save(ArtifactsDir + "FontSettings.FallbackSettings.Default.xml");

            // Below are two pre-defined font fallback schemes we can choose from.
            // 1 -  Use the default Microsoft Office scheme, which is the same one as the default:
            fontFallbackSettings.LoadMsOfficeFallbackSettings();
            fontFallbackSettings.Save(ArtifactsDir + "FontSettings.FallbackSettings.LoadMsOfficeFallbackSettings.xml");

            // 2 -  Use the scheme built from Google Noto fonts:
            fontFallbackSettings.LoadNotoFallbackSettings();
            fontFallbackSettings.Save(ArtifactsDir + "FontSettings.FallbackSettings.LoadNotoFallbackSettings.xml");
            //ExEnd

            XmlDocument fallbackSettingsDoc = new XmlDocument();
            fallbackSettingsDoc.LoadXml(File.ReadAllText(ArtifactsDir + "FontSettings.FallbackSettings.Default.xml"));
            XmlNamespaceManager manager = new XmlNamespaceManager(fallbackSettingsDoc.NameTable);
            manager.AddNamespace("aw", "Aspose.Words");

            XmlNodeList rules =
                fallbackSettingsDoc.SelectNodes("//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);

            Assert.That(rules[9].Attributes["Ranges"].Value, Is.EqualTo("0C00-0C7F"));
            Assert.That(rules[9].Attributes["FallbackFonts"].Value, Is.EqualTo("Vani"));
        }

        [Test]
        public void FallbackSettingsCustom()
        {
            //ExStart
            //ExFor:FontSettings.FallbackSettings
            //ExFor:FontFallbackSettings
            //ExFor:FontFallbackSettings.BuildAutomatic
            //ExSummary:Shows how to distribute fallback fonts across Unicode character code ranges.
            Document doc = new Document();

            FontSettings fontSettings = new FontSettings();
            doc.FontSettings = fontSettings;
            FontFallbackSettings fontFallbackSettings = fontSettings.FallbackSettings;

            // Configure our font settings to source fonts only from the "MyFonts" folder.
            FolderFontSource folderFontSource = new FolderFontSource(FontsDir, false);
            fontSettings.SetFontsSources(new FontSourceBase[] {folderFontSource});

            // Calling the "BuildAutomatic" method will generate a fallback scheme that
            // distributes accessible fonts across as many Unicode character codes as possible.
            // In our case, it only has access to the handful of fonts inside the "MyFonts" folder.
            fontFallbackSettings.BuildAutomatic();
            fontFallbackSettings.Save(ArtifactsDir + "FontSettings.FallbackSettingsCustom.BuildAutomatic.xml");

            // We can also load a custom substitution scheme from a file like this.
            // This scheme applies the "AllegroOpen" font across the "0000-00ff" Unicode blocks, the "AllegroOpen" font across "0100-024f",
            // and the "M+ 2m" font in all other ranges that other fonts in the scheme do not cover.
            fontFallbackSettings.Load(MyDir + "Custom font fallback settings.xml");

            // Create a document builder and set its font to one that does not exist in any of our sources.
            // Our font settings will invoke the fallback scheme for characters that we type using the unavailable font.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Font.Name = "Missing Font";

            // Use the builder to print every Unicode character from 0x0021 to 0x052F,
            // with descriptive lines dividing Unicode blocks we defined in our custom font fallback scheme.
            for (int i = 0x0021; i < 0x0530; i++)
            {
                switch (i)
                {
                    case 0x0021:
                        builder.Writeln(
                            "\n\n0x0021 - 0x00FF: \nBasic Latin/Latin-1 Supplement Unicode blocks in \"AllegroOpen\" font:");
                        break;
                    case 0x0100:
                        builder.Writeln(
                            "\n\n0x0100 - 0x024F: \nLatin Extended A/B blocks, mostly in \"AllegroOpen\" font:");
                        break;
                    case 0x0250:
                        builder.Writeln("\n\n0x0250 - 0x052F: \nIPA/Greek/Cyrillic blocks in \"M+ 2m\" font:");
                        break;
                }

                builder.Write($"{Convert.ToChar(i)}");
            }

            doc.Save(ArtifactsDir + "FontSettings.FallbackSettingsCustom.pdf");
            //ExEnd

            XmlDocument fallbackSettingsDoc = new XmlDocument();
            fallbackSettingsDoc.LoadXml(
                File.ReadAllText(ArtifactsDir + "FontSettings.FallbackSettingsCustom.BuildAutomatic.xml"));
            XmlNamespaceManager manager = new XmlNamespaceManager(fallbackSettingsDoc.NameTable);
            manager.AddNamespace("aw", "Aspose.Words");

            XmlNodeList rules =
                fallbackSettingsDoc.SelectNodes("//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);

            Assert.That(rules[0].Attributes["Ranges"].Value, Is.EqualTo("0000-007F"));
            Assert.That(rules[0].Attributes["FallbackFonts"].Value, Is.EqualTo("AllegroOpen"));

            Assert.That(rules[2].Attributes["Ranges"].Value, Is.EqualTo("0100-017F"));
            Assert.That(rules[2].Attributes["FallbackFonts"].Value, Is.EqualTo("AllegroOpen"));

            Assert.That(rules[4].Attributes["Ranges"].Value, Is.EqualTo("0250-02AF"));
            Assert.That(rules[4].Attributes["FallbackFonts"].Value, Is.EqualTo("M+ 2m"));

            Assert.That(rules[7].Attributes["Ranges"].Value, Is.EqualTo("0370-03FF"));
            Assert.That(rules[7].Attributes["FallbackFonts"].Value, Is.EqualTo("Arvo"));
        }

        [Test]
        public void TableSubstitutionRule()
        {
            //ExStart
            //ExFor:TableSubstitutionRule
            //ExFor:TableSubstitutionRule.LoadLinuxSettings
            //ExFor:TableSubstitutionRule.LoadWindowsSettings
            //ExFor:TableSubstitutionRule.Save(Stream)
            //ExFor:TableSubstitutionRule.Save(String)
            //ExSummary:Shows how to access font substitution tables for Windows and Linux.
            Document doc = new Document();
            FontSettings fontSettings = new FontSettings();
            doc.FontSettings = fontSettings;

            // Create a new table substitution rule and load the default Microsoft Windows font substitution table.
            TableSubstitutionRule tableSubstitutionRule = fontSettings.SubstitutionSettings.TableSubstitution;
            tableSubstitutionRule.LoadWindowsSettings();

            // In Windows, the default substitute for the "Times New Roman CE" font is "Times New Roman".
            Assert.That(tableSubstitutionRule.GetSubstitutes("Times New Roman CE").ToArray(), Is.EqualTo(new[] {"Times New Roman"}));

            // We can save the table in the form of an XML document.
            tableSubstitutionRule.Save(ArtifactsDir + "FontSettings.TableSubstitutionRule.Windows.xml");

            // Linux has its own substitution table.
            // There are multiple substitute fonts for "Times New Roman CE".
            // If the first substitute, "FreeSerif" is also unavailable,
            // this rule will cycle through the others in the array until it finds an available one.
            tableSubstitutionRule.LoadLinuxSettings();
            Assert.That(tableSubstitutionRule.GetSubstitutes("Times New Roman CE").ToArray(), Is.EqualTo(new[] {"FreeSerif", "Liberation Serif", "DejaVu Serif"}));

            // Save the Linux substitution table in the form of an XML document using a stream.
            using (FileStream fileStream = new FileStream(ArtifactsDir + "FontSettings.TableSubstitutionRule.Linux.xml",
                FileMode.Create))
            {
                tableSubstitutionRule.Save(fileStream);
            }
            //ExEnd

            XmlDocument fallbackSettingsDoc = new XmlDocument();
            fallbackSettingsDoc.LoadXml(
                File.ReadAllText(ArtifactsDir + "FontSettings.TableSubstitutionRule.Windows.xml"));
            XmlNamespaceManager manager = new XmlNamespaceManager(fallbackSettingsDoc.NameTable);
            manager.AddNamespace("aw", "Aspose.Words");

            XmlNodeList rules =
                fallbackSettingsDoc.SelectNodes("//aw:TableSubstitutionSettings/aw:SubstitutesTable/aw:Item", manager);

            Assert.That(rules[16].Attributes["OriginalFont"].Value, Is.EqualTo("Times New Roman CE"));
            Assert.That(rules[16].Attributes["SubstituteFonts"].Value, Is.EqualTo("Times New Roman"));

            fallbackSettingsDoc.LoadXml(
                File.ReadAllText(ArtifactsDir + "FontSettings.TableSubstitutionRule.Linux.xml"));
            rules = fallbackSettingsDoc.SelectNodes("//aw:TableSubstitutionSettings/aw:SubstitutesTable/aw:Item",
                manager);

            Assert.That(rules[31].Attributes["OriginalFont"].Value, Is.EqualTo("Times New Roman CE"));
            Assert.That(rules[31].Attributes["SubstituteFonts"].Value, Is.EqualTo("FreeSerif, Liberation Serif, DejaVu Serif"));
        }

        [Test]
        public void TableSubstitutionRuleCustom()
        {
            //ExStart
            //ExFor:FontSubstitutionSettings.TableSubstitution
            //ExFor:TableSubstitutionRule.AddSubstitutes(String,String[])
            //ExFor:TableSubstitutionRule.GetSubstitutes(String)
            //ExFor:TableSubstitutionRule.Load(Stream)
            //ExFor:TableSubstitutionRule.Load(String)
            //ExFor:TableSubstitutionRule.SetSubstitutes(String,String[])
            //ExSummary:Shows how to work with custom font substitution tables.
            Document doc = new Document();
            FontSettings fontSettings = new FontSettings();
            doc.FontSettings = fontSettings;

            // Create a new table substitution rule and load the default Windows font substitution table.
            TableSubstitutionRule tableSubstitutionRule = fontSettings.SubstitutionSettings.TableSubstitution;

            // If we select fonts exclusively from our folder, we will need a custom substitution table.
            // We will no longer have access to the Microsoft Windows fonts,
            // such as "Arial" or "Times New Roman" since they do not exist in our new font folder.
            FolderFontSource folderFontSource = new FolderFontSource(FontsDir, false);
            fontSettings.SetFontsSources(new FontSourceBase[] {folderFontSource});

            // Below are two ways of loading a substitution table from a file in the local file system.
            // 1 -  From a stream:
            using (FileStream fileStream = new FileStream(MyDir + "Font substitution rules.xml", FileMode.Open))
            {
                tableSubstitutionRule.Load(fileStream);
            }

            // 2 -  Directly from a file:
            tableSubstitutionRule.Load(MyDir + "Font substitution rules.xml");

            // Since we no longer have access to "Arial", our font table will first try substitute it with "Nonexistent Font".
            // We do not have this font so that it will move onto the next substitute, "Kreon", found in the "MyFonts" folder.
            Assert.That(tableSubstitutionRule.GetSubstitutes("Arial").ToArray(), Is.EqualTo(new[] {"Missing Font", "Kreon"}));

            // We can expand this table programmatically. We will add an entry that substitutes "Times New Roman" with "Arvo"
            Assert.That(tableSubstitutionRule.GetSubstitutes("Times New Roman"), Is.Null);
            tableSubstitutionRule.AddSubstitutes("Times New Roman", "Arvo");
            Assert.That(tableSubstitutionRule.GetSubstitutes("Times New Roman").ToArray(), Is.EqualTo(new[] {"Arvo"}));

            // We can add a secondary fallback substitute for an existing font entry with AddSubstitutes().
            // In case "Arvo" is unavailable, our table will look for "M+ 2m" as a second substitute option.
            tableSubstitutionRule.AddSubstitutes("Times New Roman", "M+ 2m");
            Assert.That(tableSubstitutionRule.GetSubstitutes("Times New Roman").ToArray(), Is.EqualTo(new[] {"Arvo", "M+ 2m"}));

            // SetSubstitutes() can set a new list of substitute fonts for a font.
            tableSubstitutionRule.SetSubstitutes("Times New Roman", "Squarish Sans CT", "M+ 2m");
            Assert.That(tableSubstitutionRule.GetSubstitutes("Times New Roman").ToArray(), Is.EqualTo(new[] {"Squarish Sans CT", "M+ 2m"}));

            // Writing text in fonts that we do not have access to will invoke our substitution rules.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Font.Name = "Arial";
            builder.Writeln("Text written in Arial, to be substituted by Kreon.");

            builder.Font.Name = "Times New Roman";
            builder.Writeln("Text written in Times New Roman, to be substituted by Squarish Sans CT.");

            doc.Save(ArtifactsDir + "FontSettings.TableSubstitutionRule.Custom.pdf");
            //ExEnd
        }

        [Test]
        public void ResolveFontsBeforeLoadingDocument()
        {
            //ExStart
            //ExFor:LoadOptions.FontSettings
            //ExSummary:Shows how to designate font substitutes during loading.
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.FontSettings = new FontSettings();

            // Set a font substitution rule for a LoadOptions object.
            // If the document we are loading uses a font which we do not have,
            // this rule will substitute the unavailable font with one that does exist.
            // In this case, all uses of the "MissingFont" will convert to "Comic Sans MS".
            TableSubstitutionRule substitutionRule = loadOptions.FontSettings.SubstitutionSettings.TableSubstitution;
            substitutionRule.AddSubstitutes("MissingFont", "Comic Sans MS");

            Document doc = new Document(MyDir + "Missing font.html", loadOptions);

            // At this point such text will still be in "MissingFont".
            // Font substitution will take place when we render the document.
            Assert.That(doc.FirstSection.Body.FirstParagraph.Runs[0].Font.Name, Is.EqualTo("MissingFont"));

            doc.Save(ArtifactsDir + "FontSettings.ResolveFontsBeforeLoadingDocument.pdf");
            //ExEnd
        }

        //ExStart
        //ExFor:StreamFontSource
        //ExFor:StreamFontSource.OpenFontDataStream
        //ExSummary:Shows how to load fonts from stream.
        [Test] //ExSkip
        public void StreamFontSourceFileRendering()
        {
            FontSettings fontSettings = new FontSettings();
            fontSettings.SetFontsSources(new FontSourceBase[] {new StreamFontSourceFile()});

            DocumentBuilder builder = new DocumentBuilder();
            builder.Document.FontSettings = fontSettings;
            builder.Font.Name = "Kreon-Regular";
            builder.Writeln("Test aspose text when saving to PDF.");

            builder.Document.Save(ArtifactsDir + "FontSettings.StreamFontSourceFileRendering.pdf");
        }

        /// <summary>
        /// Load the font data only when required instead of storing it in the memory
        /// for the entire lifetime of the "FontSettings" object.
        /// </summary>
        private class StreamFontSourceFile : StreamFontSource
        {
            public override Stream OpenFontDataStream()
            {
                return File.OpenRead(FontsDir + "Kreon-Regular.ttf");
            }
        }
        //ExEnd

        //ExStart
        //ExFor:FileFontSource.#ctor(String, Int32, String)
        //ExFor:MemoryFontSource.#ctor(Byte[], Int32, String)
        //ExFor:FontSettings.SaveSearchCache(Stream)
        //ExFor:FontSettings.SetFontsSources(FontSourceBase[], Stream)
        //ExFor:FileFontSource.CacheKey
        //ExFor:MemoryFontSource.CacheKey
        //ExFor:StreamFontSource.CacheKey
        //ExSummary:Shows how to speed up the font cache initialization process.
        [Test]//ExSkip
        public void LoadFontSearchCache()
        {
            const string cacheKey1 = "Arvo";
            const string cacheKey2 = "Arvo-Bold";
            FontSettings parsedFonts = new FontSettings();
            FontSettings loadedCache = new FontSettings();

            parsedFonts.SetFontsSources(new FontSourceBase[]
            {
                new FileFontSource(FontsDir + "Arvo-Regular.ttf", 0, cacheKey1),
                new FileFontSource(FontsDir + "Arvo-Bold.ttf", 0, cacheKey2)
            });
            
            using (MemoryStream cacheStream = new MemoryStream())
            {
                parsedFonts.SaveSearchCache(cacheStream);
                loadedCache.SetFontsSources(new FontSourceBase[]
                {
                    new SearchCacheStream(cacheKey1),
                    new MemoryFontSource(File.ReadAllBytes(FontsDir + "Arvo-Bold.ttf"), 0, cacheKey2)
                }, cacheStream);
            }

            Assert.That(loadedCache.GetFontsSources().Length, Is.EqualTo(parsedFonts.GetFontsSources().Length));
        }

        /// <summary>
        /// Load the font data only when required instead of storing it in the memory
        /// for the entire lifetime of the "FontSettings" object.
        /// </summary>
        private class SearchCacheStream : StreamFontSource
        {
            public SearchCacheStream(string cacheKey):base(0, cacheKey)
            {
            }

            public override Stream OpenFontDataStream()
            {
                return File.OpenRead(FontsDir + "Arvo-Regular.ttf");
            }
        }
        //ExEnd
    }
}
