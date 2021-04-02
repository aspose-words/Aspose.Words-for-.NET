// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

#if NET462 || NETCOREAPP2_1 || JAVA
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Loading;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    class ExFontSettings : ApiExampleBase
    {
        [Test]
        public void DefaultFontInstance()
        {
            //ExStart
            //ExFor:Fonts.FontSettings.DefaultInstance
            //ExSummary:Shows how to configure the default font settings instance.
            // Configure the default font settings instance to use the "Courier New" font
            // as a backup substitute when we attempt to use an unknown font.
            FontSettings.DefaultInstance.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Courier New";

            Assert.True(FontSettings.DefaultInstance.SubstitutionSettings.DefaultFontSubstitution.Enabled);

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Non-existent font";
            builder.Write("Hello world!");

            // This document does not have a FontSettings configuration. When we render the document,
            // the default FontSettings instance will resolve the missing font.
            // Aspose.Words will use "Courier New" to render text that uses the unknown font.
            Assert.Null(doc.FontSettings);

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
            Assert.AreEqual(1, fontSources.Length);
            Assert.True(fontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arial"));
            Assert.False(fontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arvo"));

            // Set the "DefaultFontName" property to "Courier New" to,
            // while rendering the document, apply that font in all cases when another font is not available. 
            FontSettings.DefaultInstance.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Courier New";

            Assert.True(fontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Courier New"));

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

            Assert.That(callback.FontWarnings.Count, Is.GreaterThan(0));
            Assert.True(callback.FontWarnings[0].WarningType == WarningType.FontSubstitution);
            Assert.True(callback.FontWarnings[0].Description.Contains("has not been found"));

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
        //ExFor:Fonts.FontSettings.DefaultInstance
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

            Assert.AreEqual(1, callback.FontSubstitutionWarnings.Count); //ExSkip
            Assert.True(callback.FontSubstitutionWarnings[0].WarningType == WarningType.FontSubstitution);
            Assert.True(callback.FontSubstitutionWarnings[0].Description
                .Equals("Font 'Times New Roman' has not been found. Using 'Fanwood' font instead. Reason: first available font."));
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
        [Test]
        public void FontSourceWarning()
        {
            FontSettings settings = new FontSettings();
            settings.SetFontsFolder("bad folder?", false);

            FontSourceBase source = settings.GetFontsSources()[0];
            FontSourceWarningCollector callback = new FontSourceWarningCollector();
            source.WarningCallback = callback;

            // Get the list of fonts to call warning callback.
            IList<PhysicalFontInfo> fontInfos = source.GetAvailableFonts();

            Assert.AreEqual("Error loading font from the folder \"bad folder?\": Illegal characters in path.",
                callback.FontSubstitutionWarnings[0].Description);
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

        //ExStart
        //ExFor:Fonts.FontInfoSubstitutionRule
        //ExFor:Fonts.FontSubstitutionSettings.FontInfoSubstitution
        //ExFor:IWarningCallback
        //ExFor:IWarningCallback.Warning(WarningInfo)
        //ExFor:WarningInfo
        //ExFor:WarningInfo.Description
        //ExFor:WarningInfo.WarningType
        //ExFor:WarningInfoCollection
        //ExFor:WarningInfoCollection.Warning(WarningInfo)
        //ExFor:WarningInfoCollection.GetEnumerator
        //ExFor:WarningInfoCollection.Clear
        //ExFor:WarningType
        //ExFor:DocumentBase.WarningCallback
        //ExSummary:Shows how to set the property for finding the closest match for a missing font from the available font sources.
        [Test]
        public void EnableFontSubstitution()
        {
            // Open a document that contains text formatted with a font that does not exist in any of our font sources.
            Document doc = new Document(MyDir + "Missing font.docx");

            // Assign a callback for handling font substitution warnings.
            HandleDocumentSubstitutionWarnings substitutionWarningHandler = new HandleDocumentSubstitutionWarnings();
            doc.WarningCallback = substitutionWarningHandler;

            // Set a default font name and enable font substitution.
            FontSettings fontSettings = new FontSettings();
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial"; ;
            fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = true;

            // We will get a font substitution warning if we save a document with a missing font.
            doc.FontSettings = fontSettings;
            doc.Save(ArtifactsDir + "FontSettings.EnableFontSubstitution.pdf");

            using (IEnumerator<WarningInfo> warnings = substitutionWarningHandler.FontWarnings.GetEnumerator())
                while (warnings.MoveNext())
                    Console.WriteLine(warnings.Current.Description);

            // We can also verify warnings in the collection and clear them.
            Assert.AreEqual(WarningSource.Layout, substitutionWarningHandler.FontWarnings[0].Source);
            Assert.AreEqual("Font '28 Days Later' has not been found. Using 'Calibri' font instead. Reason: alternative name from document.",
                substitutionWarningHandler.FontWarnings[0].Description);

            substitutionWarningHandler.FontWarnings.Clear();

            Assert.That(substitutionWarningHandler.FontWarnings, Is.Empty);
        }

        public class HandleDocumentSubstitutionWarnings : IWarningCallback
        {
            /// <summary>
            /// Called every time a warning occurs during loading/saving.
            /// </summary>
            public void Warning(WarningInfo info)
            {
                if (info.WarningType == WarningType.FontSubstitution)
                    FontWarnings.Warning(info);
            }

            public WarningInfoCollection FontWarnings = new WarningInfoCollection();
        }
        //ExEnd

        [Test]
        public void SubstitutionWarningsClosestMatch()
        {
            Document doc = new Document(MyDir + "Bullet points with alternative font.docx");

            HandleDocumentSubstitutionWarnings callback = new HandleDocumentSubstitutionWarnings();
            doc.WarningCallback = callback;

            doc.Save(ArtifactsDir + "FontSettings.SubstitutionWarningsClosestMatch.pdf");

            Assert.True(callback.FontWarnings[0].Description
                .Equals("Font \'SymbolPS\' has not been found. Using \'Wingdings\' font instead. Reason: font info substitution."));
        }

        [Test]
        public void DisableFontSubstitution()
        {
            Document doc = new Document(MyDir + "Missing font.docx");

            HandleDocumentSubstitutionWarnings callback = new HandleDocumentSubstitutionWarnings();
            doc.WarningCallback = callback;

            FontSettings fontSettings = new FontSettings();
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
            fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = false;

            doc.FontSettings = fontSettings;
            doc.Save(ArtifactsDir + "FontSettings.DisableFontSubstitution.pdf");

            Regex reg = new Regex("Font '28 Days Later' has not been found. Using (.*) font instead. Reason: default font setting.");

            foreach (WarningInfo fontWarning in callback.FontWarnings)
            {
                Match match = reg.Match(fontWarning.Description);
                if (match.Success)
                {
                    Assert.Pass();
                    break;
                }
            }
        }

        [Test]
        [Category("SkipMono")]
        public void SubstitutionWarnings()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            HandleDocumentSubstitutionWarnings callback = new HandleDocumentSubstitutionWarnings();
            doc.WarningCallback = callback;

            FontSettings fontSettings = new FontSettings();
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
            fontSettings.SetFontsFolder(FontsDir, false);
            fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Arial", "Arvo", "Slab");

            doc.FontSettings = fontSettings;
            doc.Save(ArtifactsDir + "FontSettings.SubstitutionWarnings.pdf");

            Assert.AreEqual("Font \'Arial\' has not been found. Using \'Arvo\' font instead. Reason: table substitution.",
                callback.FontWarnings[0].Description);
            Assert.AreEqual("Font \'Times New Roman\' has not been found. Using \'M+ 2m\' font instead. Reason: font info substitution.",
                callback.FontWarnings[1].Description);
        }

        [Test]
        public void GetSubstitutionWithoutSuffixes()
        {
            Document doc = new Document(MyDir + "Get substitution without suffixes.docx");

            FontSourceBase[] originalFontSources = FontSettings.DefaultInstance.GetFontsSources();

            HandleDocumentSubstitutionWarnings substitutionWarningHandler = new HandleDocumentSubstitutionWarnings();
            doc.WarningCallback = substitutionWarningHandler;

            List<FontSourceBase> fontSources = new List<FontSourceBase>(FontSettings.DefaultInstance.GetFontsSources());
            FolderFontSource folderFontSource = new FolderFontSource(FontsDir, true);
            fontSources.Add(folderFontSource);

            FontSourceBase[] updatedFontSources = fontSources.ToArray();
            FontSettings.DefaultInstance.SetFontsSources(updatedFontSources);

            doc.Save(ArtifactsDir + "Font.GetSubstitutionWithoutSuffixes.pdf");

            Assert.AreEqual(
                "Font 'DINOT-Regular' has not been found. Using 'DINOT' font instead. Reason: font name substitution.",
                substitutionWarningHandler.FontWarnings[0].Description);

            FontSettings.DefaultInstance.SetFontsSources(originalFontSources);
        }

        [Test]
        public void FontSourceFile()
        {
            //ExStart
            //ExFor:Fonts.FileFontSource
            //ExFor:Fonts.FileFontSource.#ctor(String)
            //ExFor:Fonts.FileFontSource.#ctor(String, Int32)
            //ExFor:Fonts.FileFontSource.FilePath
            //ExFor:Fonts.FileFontSource.Type
            //ExFor:Fonts.FontSourceBase
            //ExFor:Fonts.FontSourceBase.Priority
            //ExFor:Fonts.FontSourceBase.Type
            //ExFor:Fonts.FontSourceType
            //ExSummary:Shows how to use a font file in the local file system as a font source.
            FileFontSource fileFontSource = new FileFontSource(MyDir + "Alte DIN 1451 Mittelschrift.ttf", 0);

            Document doc = new Document();
            doc.FontSettings = new FontSettings();
            doc.FontSettings.SetFontsSources(new FontSourceBase[] { fileFontSource });

            Assert.AreEqual(MyDir + "Alte DIN 1451 Mittelschrift.ttf", fileFontSource.FilePath);
            Assert.AreEqual(FontSourceType.FontFile, fileFontSource.Type);
            Assert.AreEqual(0, fileFontSource.Priority);
            //ExEnd
        }

        [Test]
        public void FontSourceFolder()
        {
            //ExStart
            //ExFor:Fonts.FolderFontSource
            //ExFor:Fonts.FolderFontSource.#ctor(String, Boolean)
            //ExFor:Fonts.FolderFontSource.#ctor(String, Boolean, Int32)
            //ExFor:Fonts.FolderFontSource.FolderPath
            //ExFor:Fonts.FolderFontSource.ScanSubfolders
            //ExFor:Fonts.FolderFontSource.Type
            //ExSummary:Shows how to use a local system folder which contains fonts as a font source.

            // Create a font source from a folder that contains font files.
            FolderFontSource folderFontSource = new FolderFontSource(FontsDir, false, 1);

            Document doc = new Document();
            doc.FontSettings = new FontSettings();
            doc.FontSettings.SetFontsSources(new FontSourceBase[] { folderFontSource });

            Assert.AreEqual(FontsDir, folderFontSource.FolderPath);
            Assert.AreEqual(false, folderFontSource.ScanSubfolders);
            Assert.AreEqual(FontSourceType.FontsFolder, folderFontSource.Type);
            Assert.AreEqual(1, folderFontSource.Priority);
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

            Assert.AreEqual(1, originalFontSources.Length);
            Assert.True(originalFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arial"));

            // The default font sources are missing the two fonts that we are using in this document.
            Assert.False(originalFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arvo"));
            Assert.False(originalFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Amethysta"));

            // Use the "SetFontsFolder" method to set a directory which will act as a new font source.
            // Pass "false" as the "recursive" argument to include fonts from all the font files that are in the directory
            // that we are passing in the first argument, but not include any fonts in any of that directory's subfolders.
            // Pass "true" as the "recursive" argument to include all font files in the directory that we are passing
            // in the first argument, as well as all the fonts in its subdirectories.
            FontSettings.DefaultInstance.SetFontsFolder(FontsDir, recursive);

            FontSourceBase[] newFontSources = FontSettings.DefaultInstance.GetFontsSources();

            Assert.AreEqual(1, newFontSources.Length);
            Assert.False(newFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arial"));
            Assert.True(newFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arvo"));

            // The "Amethysta" font is in a subfolder of the font directory.
            if (recursive)
            {
                Assert.AreEqual(25, newFontSources[0].GetAvailableFonts().Count);
                Assert.True(newFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Amethysta"));
            }
            else
            {
                Assert.AreEqual(18, newFontSources[0].GetAvailableFonts().Count);
                Assert.False(newFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Amethysta"));
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

            Assert.AreEqual(1, originalFontSources.Length);
            Assert.True(originalFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arial"));

            // The default font sources are missing the two fonts that we are using in this document.
            Assert.False(originalFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Amethysta"));
            Assert.False(originalFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Junction Light"));

            // Use the "SetFontsFolders" method to create a font source from each font directory that we pass as the first argument.
            // Pass "false" as the "recursive" argument to include fonts from all the font files that are in the directories
            // that we are passing in the first argument, but not include any fonts from any of the directories' subfolders.
            // Pass "true" as the "recursive" argument to include all font files in the directories that we are passing
            // in the first argument, as well as all the fonts in their subdirectories.
            FontSettings.DefaultInstance.SetFontsFolders(new[] { FontsDir + "/Amethysta", FontsDir + "/Junction" }, recursive);

            FontSourceBase[] newFontSources = FontSettings.DefaultInstance.GetFontsSources();

            Assert.AreEqual(2, newFontSources.Length);
            Assert.False(newFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arial"));
            Assert.AreEqual(1, newFontSources[0].GetAvailableFonts().Count);
            Assert.True(newFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Amethysta"));

            // The "Junction" folder itself contains no font files, but has subfolders that do.
            if (recursive)
            {
                Assert.AreEqual(6, newFontSources[1].GetAvailableFonts().Count);
                Assert.True(newFontSources[1].GetAvailableFonts().Any(f => f.FullFontName == "Junction Light"));
            }
            else
            {
                Assert.AreEqual(0, newFontSources[1].GetAvailableFonts().Count);
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
            //ExFor:FontSettings.SetFontsSources()
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

            Assert.AreEqual(1, originalFontSources.Length);

            Assert.True(originalFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arial"));

            // The default font source is missing two of the fonts that we are using in our document.
            // When we save this document, Aspose.Words will apply fallback fonts to all text formatted with inaccessible fonts.
            Assert.False(originalFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Amethysta"));
            Assert.False(originalFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Junction Light"));

            // Create a font source from a folder that contains fonts.
            FolderFontSource folderFontSource = new FolderFontSource(FontsDir, true);

            // Apply a new array of font sources that contains the original font sources, as well as our custom fonts.
            FontSourceBase[] updatedFontSources = { originalFontSources[0], folderFontSource };
            FontSettings.DefaultInstance.SetFontsSources(updatedFontSources);

            // Verify that Aspose.Words has access to all required fonts before we render the document to PDF.
            updatedFontSources = FontSettings.DefaultInstance.GetFontsSources();

            Assert.True(updatedFontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arial"));
            Assert.True(updatedFontSources[1].GetAvailableFonts().Any(f => f.FullFontName == "Amethysta"));
            Assert.True(updatedFontSources[1].GetAvailableFonts().Any(f => f.FullFontName == "Junction Light"));

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

            FolderFontSource folderSource = ((FolderFontSource)doc.FontSettings.GetFontsSources()[0]);

            Assert.AreEqual(FontsDir, folderSource.FolderPath);
            Assert.False(folderSource.ScanSubfolders);
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
            Assert.AreEqual(1, fontSources.Length);
            Assert.True(fontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arial"));

            // The second font, "Amethysta", is unavailable.
            Assert.False(fontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Amethysta"));

            // We can configure a font substitution table which determines
            // which fonts Aspose.Words will use as substitutes for unavailable fonts.
            // Set two substitution fonts for "Amethysta": "Arvo", and "Courier New".
            // If the first substitute is unavailable, Aspose.Words attempts to use the second substitute, and so on.
            doc.FontSettings = new FontSettings();
            doc.FontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes(
                "Amethysta", new[] { "Arvo", "Courier New" });

            // "Amethysta" is unavailable, and the substitution rule states that the first font to use as a substitute is "Arvo". 
            Assert.False(fontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Arvo"));

            // "Arvo" is also unavailable, but "Courier New" is. 
            Assert.True(fontSources[0].GetAvailableFonts().Any(f => f.FullFontName == "Courier New"));

            // The output document will display the text that uses the "Amethysta" font formatted with "Courier New".
            doc.Save(ArtifactsDir + "FontSettings.TableSubstitution.pdf");
            //ExEnd
        }

        [Test]
        public void SetSpecifyFontFolders()
        {
            FontSettings fontSettings = new FontSettings();
            fontSettings.SetFontsFolders(new string[] { FontsDir, @"C:\Windows\Fonts\" }, true);

            // Using load options
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.FontSettings = fontSettings;
            Document doc = new Document(MyDir + "Rendering.docx", loadOptions);

            FolderFontSource folderSource = ((FolderFontSource)doc.FontSettings.GetFontsSources()[0]);
            Assert.AreEqual(FontsDir, folderSource.FolderPath);
            Assert.True(folderSource.ScanSubfolders);

            folderSource = ((FolderFontSource)doc.FontSettings.GetFontsSources()[1]);
            Assert.AreEqual(@"C:\Windows\Fonts\", folderSource.FolderPath);
            Assert.True(folderSource.ScanSubfolders);
        }

        [Test]
        public void AddFontSubstitutes()
        {
            FontSettings fontSettings = new FontSettings();
            fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes("Slab", new string[] { "Times New Roman", "Arial" });
            fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Arvo", new string[] { "Open Sans", "Arial" });

            Document doc = new Document(MyDir + "Rendering.docx");
            doc.FontSettings = fontSettings;

            string[] alternativeFonts = doc.FontSettings.SubstitutionSettings.TableSubstitution.GetSubstitutes("Slab").ToArray();
            Assert.AreEqual(new string[] { "Times New Roman", "Arial" }, alternativeFonts);

            alternativeFonts = doc.FontSettings.SubstitutionSettings.TableSubstitution.GetSubstitutes("Arvo").ToArray();
            Assert.AreEqual(new string[] { "Open Sans", "Arial" }, alternativeFonts);
        }

        [Test]
        public void FontSourceMemory()
        {
            //ExStart
            //ExFor:Fonts.MemoryFontSource
            //ExFor:Fonts.MemoryFontSource.#ctor(Byte[])
            //ExFor:Fonts.MemoryFontSource.#ctor(Byte[], Int32)
            //ExFor:Fonts.MemoryFontSource.FontData
            //ExFor:Fonts.MemoryFontSource.Type
            //ExSummary:Shows how to use a byte array with data from a font file as a font source.

            byte[] fontBytes = File.ReadAllBytes(MyDir + "Alte DIN 1451 Mittelschrift.ttf");
            MemoryFontSource memoryFontSource = new MemoryFontSource(fontBytes, 0);

            Document doc = new Document();
            doc.FontSettings = new FontSettings();
            doc.FontSettings.SetFontsSources(new FontSourceBase[] { memoryFontSource });

            Assert.AreEqual(FontSourceType.MemoryFont, memoryFontSource.Type);
            Assert.AreEqual(0, memoryFontSource.Priority);
            //ExEnd
        }

        [Test]
        public void FontSourceSystem()
        {
            //ExStart
            //ExFor:TableSubstitutionRule.AddSubstitutes(String, String[])
            //ExFor:FontSubstitutionRule.Enabled
            //ExFor:TableSubstitutionRule.GetSubstitutes(String)
            //ExFor:Fonts.FontSettings.ResetFontSources
            //ExFor:Fonts.FontSettings.SubstitutionSettings
            //ExFor:Fonts.FontSubstitutionSettings
            //ExFor:Fonts.SystemFontSource
            //ExFor:Fonts.SystemFontSource.#ctor
            //ExFor:Fonts.SystemFontSource.#ctor(Int32)
            //ExFor:Fonts.SystemFontSource.GetSystemFontFolders
            //ExFor:Fonts.SystemFontSource.Type
            //ExSummary:Shows how to access a document's system font source and set font substitutes.
            Document doc = new Document();
            doc.FontSettings = new FontSettings();

            // By default, a blank document always contains a system font source.
            Assert.AreEqual(1, doc.FontSettings.GetFontsSources().Length);

            SystemFontSource systemFontSource = (SystemFontSource)doc.FontSettings.GetFontsSources()[0];
            Assert.AreEqual(FontSourceType.SystemFonts, systemFontSource.Type);
            Assert.AreEqual(0, systemFontSource.Priority);

            PlatformID pid = Environment.OSVersion.Platform;
            bool isWindows = (pid == PlatformID.Win32NT) || (pid == PlatformID.Win32S) || (pid == PlatformID.Win32Windows) || (pid == PlatformID.WinCE);
            if (isWindows)
            {
                const string fontsPath = @"C:\WINDOWS\Fonts";
                Assert.AreEqual(fontsPath.ToLower(), SystemFontSource.GetSystemFontFolders().FirstOrDefault()?.ToLower());
            }

            foreach (string systemFontFolder in SystemFontSource.GetSystemFontFolders())
            {
                Console.WriteLine(systemFontFolder);
            }

            // Set a font that exists in the Windows Fonts directory as a substitute for one that does not.
            doc.FontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = true;
            doc.FontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Kreon-Regular", new[] { "Calibri" });

            Assert.AreEqual(1, doc.FontSettings.SubstitutionSettings.TableSubstitution.GetSubstitutes("Kreon-Regular").Count());
            Assert.Contains("Calibri", doc.FontSettings.SubstitutionSettings.TableSubstitution.GetSubstitutes("Kreon-Regular").ToArray());

            // Alternatively, we could add a folder font source in which the corresponding folder contains the font.
            FolderFontSource folderFontSource = new FolderFontSource(FontsDir, false);
            doc.FontSettings.SetFontsSources(new FontSourceBase[] { systemFontSource, folderFontSource });
            Assert.AreEqual(2, doc.FontSettings.GetFontsSources().Length);

            // Resetting the font sources still leaves us with the system font source as well as our substitutes.
            doc.FontSettings.ResetFontSources();

            Assert.AreEqual(1, doc.FontSettings.GetFontsSources().Length);
            Assert.AreEqual(FontSourceType.SystemFonts, doc.FontSettings.GetFontsSources()[0].Type);
            Assert.AreEqual(1, doc.FontSettings.SubstitutionSettings.TableSubstitution.GetSubstitutes("Kreon-Regular").Count());
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

            XmlNodeList rules = fallbackSettingsDoc.SelectNodes("//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);

            Assert.AreEqual("0B80-0BFF", rules[0].Attributes["Ranges"].Value);
            Assert.AreEqual("Vijaya", rules[0].Attributes["FallbackFonts"].Value);

            Assert.AreEqual("1F300-1F64F", rules[1].Attributes["Ranges"].Value);
            Assert.AreEqual("Segoe UI Emoji, Segoe UI Symbol", rules[1].Attributes["FallbackFonts"].Value);

            Assert.AreEqual("2000-206F, 2070-209F, 20B9", rules[2].Attributes["Ranges"].Value);
            Assert.AreEqual("Arial", rules[2].Attributes["FallbackFonts"].Value);

            Assert.AreEqual("3040-309F", rules[3].Attributes["Ranges"].Value);
            Assert.AreEqual("MS Gothic", rules[3].Attributes["FallbackFonts"].Value);
            Assert.AreEqual("Times New Roman", rules[3].Attributes["BaseFonts"].Value);

            Assert.AreEqual("3040-309F", rules[4].Attributes["Ranges"].Value);
            Assert.AreEqual("MS Mincho", rules[4].Attributes["FallbackFonts"].Value);

            Assert.AreEqual("Arial Unicode MS", rules[5].Attributes["FallbackFonts"].Value);
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

            TestUtil.VerifyWebResponseStatusCode(HttpStatusCode.OK, "https://www.google.com/get/noto/#sans-lgc");
        }

        [Test]
        public void DefaultFontSubstitutionRule()
        {
            //ExStart
            //ExFor:Fonts.DefaultFontSubstitutionRule
            //ExFor:Fonts.DefaultFontSubstitutionRule.DefaultFontName
            //ExFor:Fonts.FontSubstitutionSettings.DefaultFontSubstitution
            //ExSummary:Shows how to set the default font substitution rule.
            Document doc = new Document();
            FontSettings fontSettings = new FontSettings();
            doc.FontSettings = fontSettings;

            // Get the default substitution rule within FontSettings.
            // This rule will substitute all missing fonts with "Times New Roman".
            DefaultFontSubstitutionRule defaultFontSubstitutionRule = fontSettings.SubstitutionSettings.DefaultFontSubstitution;
            Assert.True(defaultFontSubstitutionRule.Enabled);
            Assert.AreEqual("Times New Roman", defaultFontSubstitutionRule.DefaultFontName);

            // Set the default font substitute to "Courier New".
            defaultFontSubstitutionRule.DefaultFontName = "Courier New";

            // Using a document builder, add some text in a font that we do not have to see the substitution take place,
            // and then render the result in a PDF.
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Font.Name = "Missing Font";
            builder.Writeln("Line written in a missing font, which will be substituted with Courier New.");

            doc.Save(ArtifactsDir + "FontSettings.DefaultFontSubstitutionRule.pdf");
            //ExEnd

            Assert.AreEqual("Courier New", defaultFontSubstitutionRule.DefaultFontName);
        }

        [Test]
        public void FontConfigSubstitution()
        {
            //ExStart
            //ExFor:Fonts.FontConfigSubstitutionRule
            //ExFor:Fonts.FontConfigSubstitutionRule.Enabled
            //ExFor:Fonts.FontConfigSubstitutionRule.IsFontConfigAvailable
            //ExFor:Fonts.FontConfigSubstitutionRule.ResetCache
            //ExFor:Fonts.FontSubstitutionRule
            //ExFor:Fonts.FontSubstitutionRule.Enabled
            //ExFor:Fonts.FontSubstitutionSettings.FontConfigSubstitution
            //ExSummary:Shows operating system-dependent font config substitution.
            FontSettings fontSettings = new FontSettings();
            FontConfigSubstitutionRule fontConfigSubstitution = fontSettings.SubstitutionSettings.FontConfigSubstitution;

            bool isWindows = new[] { PlatformID.Win32NT, PlatformID.Win32S, PlatformID.Win32Windows, PlatformID.WinCE }
                .Any(p => Environment.OSVersion.Platform == p);

            // The FontConfigSubstitutionRule object works differently on Windows/non-Windows platforms.
            // On Windows, it is unavailable.
            if (isWindows)
            {
                Assert.False(fontConfigSubstitution.Enabled);
                Assert.False(fontConfigSubstitution.IsFontConfigAvailable());
            }

            bool isLinuxOrMac = new[] { PlatformID.Unix, PlatformID.MacOSX }.Any(p => Environment.OSVersion.Platform == p);

            // On Linux/Mac, we will have access to it, and will be able to perform operations.
            if (isLinuxOrMac)
            {
                Assert.True(fontConfigSubstitution.Enabled);
                Assert.True(fontConfigSubstitution.IsFontConfigAvailable());

                fontConfigSubstitution.ResetCache();
            }
            //ExEnd
        }

        [Test]
        public void FallbackSettings()
        {
            //ExStart
            //ExFor:Fonts.FontFallbackSettings.LoadMsOfficeFallbackSettings
            //ExFor:Fonts.FontFallbackSettings.LoadNotoFallbackSettings
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

            XmlNodeList rules = fallbackSettingsDoc.SelectNodes("//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);

            Assert.AreEqual("0C00-0C7F", rules[5].Attributes["Ranges"].Value);
            Assert.AreEqual("Vani", rules[5].Attributes["FallbackFonts"].Value);
        }

        [Test]
        public void FallbackSettingsCustom()
        {
            //ExStart
            //ExFor:Fonts.FontSettings.FallbackSettings
            //ExFor:Fonts.FontFallbackSettings
            //ExFor:Fonts.FontFallbackSettings.BuildAutomatic
            //ExSummary:Shows how to distribute fallback fonts across Unicode character code ranges.
            Document doc = new Document();

            FontSettings fontSettings = new FontSettings();
            doc.FontSettings = fontSettings;
            FontFallbackSettings fontFallbackSettings = fontSettings.FallbackSettings;

            // Configure our font settings to source fonts only from the "MyFonts" folder.
            FolderFontSource folderFontSource = new FolderFontSource(FontsDir, false);
            fontSettings.SetFontsSources(new FontSourceBase[] { folderFontSource });

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
                        builder.Writeln("\n\n0x0021 - 0x00FF: \nBasic Latin/Latin-1 Supplement Unicode blocks in \"AllegroOpen\" font:");
                        break;
                    case 0x0100:
                        builder.Writeln("\n\n0x0100 - 0x024F: \nLatin Extended A/B blocks, mostly in \"AllegroOpen\" font:");
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
            fallbackSettingsDoc.LoadXml(File.ReadAllText(ArtifactsDir + "FontSettings.FallbackSettingsCustom.BuildAutomatic.xml"));
            XmlNamespaceManager manager = new XmlNamespaceManager(fallbackSettingsDoc.NameTable);
            manager.AddNamespace("aw", "Aspose.Words");

            XmlNodeList rules = fallbackSettingsDoc.SelectNodes("//aw:FontFallbackSettings/aw:FallbackTable/aw:Rule", manager);

            Assert.AreEqual("0000-007F", rules[0].Attributes["Ranges"].Value);
            Assert.AreEqual("AllegroOpen", rules[0].Attributes["FallbackFonts"].Value);
            
            Assert.AreEqual("0100-017F", rules[2].Attributes["Ranges"].Value);
            Assert.AreEqual("AllegroOpen", rules[2].Attributes["FallbackFonts"].Value);

            Assert.AreEqual("0250-02AF", rules[4].Attributes["Ranges"].Value);
            Assert.AreEqual("M+ 2m", rules[4].Attributes["FallbackFonts"].Value);

            Assert.AreEqual("0370-03FF", rules[7].Attributes["Ranges"].Value);
            Assert.AreEqual("Arvo", rules[7].Attributes["FallbackFonts"].Value);
        }

        [Test]
        public void TableSubstitutionRule()
        {
            //ExStart
            //ExFor:Fonts.TableSubstitutionRule
            //ExFor:Fonts.TableSubstitutionRule.LoadLinuxSettings
            //ExFor:Fonts.TableSubstitutionRule.LoadWindowsSettings
            //ExFor:Fonts.TableSubstitutionRule.Save(System.IO.Stream)
            //ExFor:Fonts.TableSubstitutionRule.Save(System.String)
            //ExSummary:Shows how to access font substitution tables for Windows and Linux.
            Document doc = new Document();
            FontSettings fontSettings = new FontSettings();
            doc.FontSettings = fontSettings;

            // Create a new table substitution rule and load the default Microsoft Windows font substitution table.
            TableSubstitutionRule tableSubstitutionRule = fontSettings.SubstitutionSettings.TableSubstitution;
            tableSubstitutionRule.LoadWindowsSettings();

            // In Windows, the default substitute for the "Times New Roman CE" font is "Times New Roman".
            Assert.AreEqual(new[] { "Times New Roman" },
                tableSubstitutionRule.GetSubstitutes("Times New Roman CE").ToArray());

            // We can save the table in the form of an XML document.
            tableSubstitutionRule.Save(ArtifactsDir + "FontSettings.TableSubstitutionRule.Windows.xml");

            // Linux has its own substitution table.
            // There are multiple substitute fonts for "Times New Roman CE".
            // If the first substitute, "FreeSerif" is also unavailable,
            // this rule will cycle through the others in the array until it finds an available one.
            tableSubstitutionRule.LoadLinuxSettings();
            Assert.AreEqual(new[] { "FreeSerif", "Liberation Serif", "DejaVu Serif" },
                tableSubstitutionRule.GetSubstitutes("Times New Roman CE").ToArray());

            // Save the Linux substitution table in the form of an XML document using a stream.
            using (FileStream fileStream = new FileStream(ArtifactsDir + "FontSettings.TableSubstitutionRule.Linux.xml", FileMode.Create))
            {
                tableSubstitutionRule.Save(fileStream);
            }
            //ExEnd

            XmlDocument fallbackSettingsDoc = new XmlDocument();
            fallbackSettingsDoc.LoadXml(File.ReadAllText(ArtifactsDir + "FontSettings.TableSubstitutionRule.Windows.xml"));
            XmlNamespaceManager manager = new XmlNamespaceManager(fallbackSettingsDoc.NameTable);
            manager.AddNamespace("aw", "Aspose.Words");

            XmlNodeList rules = fallbackSettingsDoc.SelectNodes("//aw:TableSubstitutionSettings/aw:SubstitutesTable/aw:Item", manager);

            Assert.AreEqual("Times New Roman CE", rules[16].Attributes["OriginalFont"].Value);
            Assert.AreEqual("Times New Roman", rules[16].Attributes["SubstituteFonts"].Value);

            fallbackSettingsDoc.LoadXml(File.ReadAllText(ArtifactsDir + "FontSettings.TableSubstitutionRule.Linux.xml"));
            rules = fallbackSettingsDoc.SelectNodes("//aw:TableSubstitutionSettings/aw:SubstitutesTable/aw:Item", manager);

            Assert.AreEqual("Times New Roman CE", rules[31].Attributes["OriginalFont"].Value);
            Assert.AreEqual("FreeSerif, Liberation Serif, DejaVu Serif", rules[31].Attributes["SubstituteFonts"].Value);
        }

        [Test]
        public void TableSubstitutionRuleCustom()
        {
            //ExStart
            //ExFor:Fonts.FontSubstitutionSettings.TableSubstitution
            //ExFor:Fonts.TableSubstitutionRule.AddSubstitutes(System.String,System.String[])
            //ExFor:Fonts.TableSubstitutionRule.GetSubstitutes(System.String)
            //ExFor:Fonts.TableSubstitutionRule.Load(System.IO.Stream)
            //ExFor:Fonts.TableSubstitutionRule.Load(System.String)
            //ExFor:Fonts.TableSubstitutionRule.SetSubstitutes(System.String,System.String[])
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
            fontSettings.SetFontsSources(new FontSourceBase[] { folderFontSource });

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
            Assert.AreEqual(new[] { "Missing Font", "Kreon" }, tableSubstitutionRule.GetSubstitutes("Arial").ToArray());

            // We can expand this table programmatically. We will add an entry that substitutes "Times New Roman" with "Arvo"
            Assert.Null(tableSubstitutionRule.GetSubstitutes("Times New Roman"));
            tableSubstitutionRule.AddSubstitutes("Times New Roman", "Arvo");
            Assert.AreEqual(new[] { "Arvo" }, tableSubstitutionRule.GetSubstitutes("Times New Roman").ToArray());

            // We can add a secondary fallback substitute for an existing font entry with AddSubstitutes().
            // In case "Arvo" is unavailable, our table will look for "M+ 2m" as a second substitute option.
            tableSubstitutionRule.AddSubstitutes("Times New Roman", "M+ 2m");
            Assert.AreEqual(new[] { "Arvo", "M+ 2m" }, tableSubstitutionRule.GetSubstitutes("Times New Roman").ToArray());

            // SetSubstitutes() can set a new list of substitute fonts for a font.
            tableSubstitutionRule.SetSubstitutes("Times New Roman", new[] { "Squarish Sans CT", "M+ 2m" });
            Assert.AreEqual(new[] { "Squarish Sans CT", "M+ 2m" }, tableSubstitutionRule.GetSubstitutes("Times New Roman").ToArray());

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
            substitutionRule.AddSubstitutes("MissingFont", new[] { "Comic Sans MS" });

            Document doc = new Document(MyDir + "Missing font.html", loadOptions);

            // At this point such text will still be in "MissingFont".
            // Font substitution will take place when we render the document.
            Assert.AreEqual("MissingFont", doc.FirstSection.Body.FirstParagraph.Runs[0].Font.Name);

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
            fontSettings.SetFontsSources(new FontSourceBase[] { new StreamFontSourceFile() });

            DocumentBuilder builder = new DocumentBuilder();
            builder.Document.FontSettings = fontSettings;
            builder.Font.Name = "Kreon-Regular";
            builder.Writeln("Test aspose text when saving to PDF.");

            builder.Document.Save(ArtifactsDir + "FontSettings.StreamFontSourceFileRendering.pdf");
        }

        /// <summary>
        /// Load the font data only when required instead of storing it in the memory for the entire lifetime of the "FontSettings" object.
        /// </summary>
        private class StreamFontSourceFile : StreamFontSource
        {
            public override Stream OpenFontDataStream()
            {
                return File.OpenRead(FontsDir + "Kreon-Regular.ttf");
            }
        }
        //ExEnd
    }
}
#endif
