﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Loading;
using NUnit.Framework;
using Font = Aspose.Words.Font;

namespace DocsExamples.Programming_with_Documents
{
    internal class WorkingWithFonts : DocsExamplesBase
    {
        [Test]
        public void FontFormatting()
        {
            //ExStart:WriteAndFont
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Font font = builder.Font;
            font.Size = 16;
            font.Bold = true;
            font.Color = Color.Blue;
            font.Name = "Arial";
            font.Underline = Underline.Dash;

            builder.Write("Sample text.");
            
            doc.Save(ArtifactsDir + "WorkingWithFonts.FontFormatting.docx");
            //ExEnd:WriteAndFont
        }

        [Test]
        public void GetFontLineSpacing()
        {
            //ExStart:GetFontLineSpacing
            //GistId:7cb86f131b74afcbebc153f0039e3947
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Font.Name = "Calibri";
            builder.Writeln("qText");

            Font font = builder.Document.FirstSection.Body.FirstParagraph.Runs[0].Font;
            Console.WriteLine($"lineSpacing = {font.LineSpacing}");
            //ExEnd:GetFontLineSpacing
        }

        [Test]
        public void CheckDMLTextEffect()
        {
            //ExStart:CheckDMLTextEffect
            Document doc = new Document(MyDir + "DrawingML text effects.docx");
            
            RunCollection runs = doc.FirstSection.Body.FirstParagraph.Runs;
            Font runFont = runs[0].Font;

            // One run might have several Dml text effects applied.
            Console.WriteLine(runFont.HasDmlEffect(TextDmlEffect.Shadow));
            Console.WriteLine(runFont.HasDmlEffect(TextDmlEffect.Effect3D));
            Console.WriteLine(runFont.HasDmlEffect(TextDmlEffect.Reflection));
            Console.WriteLine(runFont.HasDmlEffect(TextDmlEffect.Outline));
            Console.WriteLine(runFont.HasDmlEffect(TextDmlEffect.Fill));
            //ExEnd:CheckDMLTextEffect
        }

        [Test]
        public void SetFontFormatting()
        {
            //ExStart:SetFontFormatting
            //GistId:7cb86f131b74afcbebc153f0039e3947
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Font font = builder.Font;
            font.Bold = true;
            font.Color = Color.DarkBlue;
            font.Italic = true;
            font.Name = "Arial";
            font.Size = 24;
            font.Spacing = 5;
            font.Underline = Underline.Double;

            builder.Writeln("I'm a very nice formatted string.");
            
            doc.Save(ArtifactsDir + "WorkingWithFonts.SetFontFormatting.docx");
            //ExEnd:SetFontFormatting
        }

        [Test]
        public void SetFontEmphasisMark()
        {
            //ExStart:SetFontEmphasisMark
            //GistId:7cb86f131b74afcbebc153f0039e3947
            Document document = new Document();
            DocumentBuilder builder = new DocumentBuilder(document);

            builder.Font.EmphasisMark = EmphasisMark.UnderSolidCircle;

            builder.Write("Emphasis text");
            builder.Writeln();
            builder.Font.ClearFormatting();
            builder.Write("Simple text");

            document.Save(ArtifactsDir + "WorkingWithFonts.SetFontEmphasisMark.docx");
            //ExEnd:SetFontEmphasisMark
        }

        [Test]
        public void EnableDisableFontSubstitution()
        {
            //ExStart:EnableDisableFontSubstitution
            Document doc = new Document(MyDir + "Rendering.docx");

            FontSettings fontSettings = new FontSettings();
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
            fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = false;
            
            doc.FontSettings = fontSettings;
            
            doc.Save(ArtifactsDir + "WorkingWithFonts.EnableDisableFontSubstitution.pdf");
            //ExEnd:EnableDisableFontSubstitution
        }

        [Test]
        public void FontFallbackSettings()
        {
            //ExStart:FontFallbackSettings
            //GistId:a08698f540d47082b4e2dbb1cb67fc1b
            Document doc = new Document(MyDir + "Rendering.docx");

            FontSettings fontSettings = new FontSettings();
            fontSettings.FallbackSettings.Load(MyDir + "Font fallback rules.xml");
            
            doc.FontSettings = fontSettings;
            
            doc.Save(ArtifactsDir + "WorkingWithFonts.FontFallbackSettings.pdf");
            //ExEnd:FontFallbackSettings
        }

        [Test]
        public void NotoFallbackSettings()
        {
            //ExStart:NotoFallbackSettings
            //GistId:a08698f540d47082b4e2dbb1cb67fc1b
            Document doc = new Document(MyDir + "Rendering.docx");

            FontSettings fontSettings = new FontSettings();
            fontSettings.FallbackSettings.LoadNotoFallbackSettings();
            
            doc.FontSettings = fontSettings;
            
            doc.Save(ArtifactsDir + "WorkingWithFonts.NotoFallbackSettings.pdf");
            //ExEnd:NotoFallbackSettings
        }

        [Test]
        public void DefaultInstance()
        {
            //ExStart:DefaultInstance
            //GistId:7e64f6d40825be58a8c12f1307c12964
            FontSettings.DefaultInstance.SetFontsFolder("C:\\MyFonts\\", true);
            //ExEnd:DefaultInstance

            Document doc = new Document(MyDir + "Rendering.docx");
            doc.Save(ArtifactsDir + "WorkingWithFonts.DefaultInstance.pdf");
        }

        [Test]
        public void MultipleFolders()
        {
            //ExStart:MultipleFolders
            //GistId:7e64f6d40825be58a8c12f1307c12964
            Document doc = new Document(MyDir + "Rendering.docx");
            
            FontSettings fontSettings = new FontSettings();
            // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
            // fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
            // FontSettings.SetFontSources instead.
            fontSettings.SetFontsFolders(new[] { @"C:\MyFonts\", @"D:\Misc\Fonts\" }, true);
            
            doc.FontSettings = fontSettings;
            
            doc.Save(ArtifactsDir + "WorkingWithFonts.MultipleFolders.pdf");
            //ExEnd:MultipleFolders
        }

        [Test]
        public void SetFontsFoldersSystemAndCustomFolder()
        {
            //ExStart:SetFontsFoldersSystemAndCustomFolder
            Document doc = new Document(MyDir + "Rendering.docx");

            FontSettings fontSettings = new FontSettings();
            // Retrieve the array of environment-dependent font sources that are searched by default.
            // For example this will contain a "Windows\Fonts\" source on a Windows machines.
            // We add this array to a new List to make adding or removing font entries much easier.
            List<FontSourceBase> fontSources = new List<FontSourceBase>(fontSettings.GetFontsSources());

            // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts.
            FolderFontSource folderFontSource = new FolderFontSource("C:\\MyFonts\\", true);
            // Add the custom folder which contains our fonts to the list of existing font sources.
            fontSources.Add(folderFontSource);

            FontSourceBase[] updatedFontSources = fontSources.ToArray();
            fontSettings.SetFontsSources(updatedFontSources);

            doc.FontSettings = fontSettings;

            doc.Save(ArtifactsDir + "WorkingWithFonts.SetFontsFoldersSystemAndCustomFolder.pdf");
            //ExEnd:SetFontsFoldersSystemAndCustomFolder
        }

        [Test]
        public void FontsFoldersWithPriority()
        {
            //ExStart:FontsFoldersWithPriority
            //GistId:7e64f6d40825be58a8c12f1307c12964
            FontSettings.DefaultInstance.SetFontsSources(new FontSourceBase[]
            {
                new SystemFontSource(), new FolderFontSource("C:\\MyFonts\\", true, 1)
            });
            //ExEnd:FontsFoldersWithPriority

            Document doc = new Document(MyDir + "Rendering.docx");
            doc.Save(ArtifactsDir + "WorkingWithFonts.FontsFoldersWithPriority.pdf");
        }

        [Test]
        public void TrueTypeFontsFolder()
        {
            //ExStart:TrueTypeFontsFolder
            //GistId:7e64f6d40825be58a8c12f1307c12964
            Document doc = new Document(MyDir + "Rendering.docx");

            FontSettings fontSettings = new FontSettings();
            // Note that this setting will override any default font sources that are being searched by default. Now only these folders will be searched for
            // Fonts when rendering or embedding fonts. To add an extra font source while keeping system font sources then use both FontSettings.GetFontSources and
            // FontSettings.SetFontSources instead
            fontSettings.SetFontsFolder(@"C:\MyFonts\", false);
            // Set font settings
            doc.FontSettings = fontSettings;
            
            doc.Save(ArtifactsDir + "WorkingWithFonts.TrueTypeFontsFolder.pdf");
            //ExEnd:TrueTypeFontsFolder
        }

        [Test]
        public void SpecifyDefaultFontWhenRendering()
        {
            //ExStart:SpecifyDefaultFontWhenRendering
            Document doc = new Document(MyDir + "Rendering.docx");

            FontSettings fontSettings = new FontSettings();
            // If the default font defined here cannot be found during rendering then
            // the closest font on the machine is used instead.
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial Unicode MS";
            
            doc.FontSettings = fontSettings;
            
            doc.Save(ArtifactsDir + "WorkingWithFonts.SpecifyDefaultFontWhenRendering.pdf");
            //ExEnd:SpecifyDefaultFontWhenRendering
        }

        [Test]
        public void FontSettingsWithLoadOptions()
        {
            //ExStart:FontSettingsWithLoadOptions
            FontSettings fontSettings = new FontSettings();

            TableSubstitutionRule substitutionRule = fontSettings.SubstitutionSettings.TableSubstitution;
            // If "UnknownFont1" font family is not available then substitute it by "Comic Sans MS"
            substitutionRule.AddSubstitutes("UnknownFont1", new string[] { "Comic Sans MS" });
            
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.FontSettings = fontSettings;
            
            Document doc = new Document(MyDir + "Rendering.docx", loadOptions);
            //ExEnd:FontSettingsWithLoadOptions
        }

        [Test]
        public void SetFontsFolder()
        {
            //ExStart:SetFontsFolder
            FontSettings fontSettings = new FontSettings();
            fontSettings.SetFontsFolder(MyDir + "Fonts", false);
            
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.FontSettings = fontSettings;
            
            Document doc = new Document(MyDir + "Rendering.docx", loadOptions);
            //ExEnd:SetFontsFolder
        }

        [Test]
        public void LoadOptionFontSettings()
        {
            //ExStart:LoadOptionFontSettings
            //GistId:a08698f540d47082b4e2dbb1cb67fc1b
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.FontSettings = new FontSettings();

            Document doc = new Document(MyDir + "Rendering.docx", loadOptions);
            //ExEnd:LoadOptionFontSettings
        }

        [Test]
        public void FontSettingsDefaultInstance()
        {
            //ExStart:FontsFolders
            //GistId:7e64f6d40825be58a8c12f1307c12964
            //ExStart:FontSettingsFontSource
            //GistId:a08698f540d47082b4e2dbb1cb67fc1b
            //ExStart:FontSettingsDefaultInstance
            //GistId:a08698f540d47082b4e2dbb1cb67fc1b
            FontSettings fontSettings = FontSettings.DefaultInstance;
            //ExEnd:FontSettingsDefaultInstance
            fontSettings.SetFontsSources(new FontSourceBase[]
            {
                new SystemFontSource(),
                new FolderFontSource("C:\\MyFonts\\", true)
            });
            //ExEnd:FontSettingsFontSource

            Document doc = new Document(MyDir + "Rendering.docx");
            //ExEnd:FontsFolders
        }

        [Test]
        public void AvailableFonts()
        {
            //ExStart:AvailableFonts
            //GistId:7e64f6d40825be58a8c12f1307c12964
            FontSettings fontSettings = new FontSettings();
            List<FontSourceBase> fontSources = new List<FontSourceBase>(fontSettings.GetFontsSources());

            // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts.
            FolderFontSource folderFontSource = new FolderFontSource(MyDir, true);
            // Add the custom folder which contains our fonts to the list of existing font sources.
            fontSources.Add(folderFontSource);

            FontSourceBase[] updatedFontSources = fontSources.ToArray();

            foreach (PhysicalFontInfo fontInfo in updatedFontSources[0].GetAvailableFonts())
            {
                Console.WriteLine("FontFamilyName : " + fontInfo.FontFamilyName);
                Console.WriteLine("FullFontName  : " + fontInfo.FullFontName);
                Console.WriteLine("Version  : " + fontInfo.Version);
                Console.WriteLine("FilePath : " + fontInfo.FilePath);
            }
            //ExEnd:AvailableFonts
        }

        [Test]
        public void ReceiveNotificationsOfFonts()
        {
            //ExStart:ReceiveNotificationsOfFonts
            Document doc = new Document(MyDir + "Rendering.docx");

            FontSettings fontSettings = new FontSettings();

            // We can choose the default font to use in the case of any missing fonts.
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";
            // For testing we will set Aspose.Words to look for fonts only in a folder which doesn't exist. Since Aspose.Words won't
            // find any fonts in the specified directory, then during rendering the fonts in the document will be subsuited with the default
            // font specified under FontSettings.DefaultFontName. We can pick up on this subsuition using our callback.
            fontSettings.SetFontsFolder(string.Empty, false);

            // Create a new class implementing IWarningCallback which collect any warnings produced during document save.
            HandleDocumentWarnings callback = new HandleDocumentWarnings();

            doc.WarningCallback = callback;
            doc.FontSettings = fontSettings;
            
            doc.Save(ArtifactsDir + "WorkingWithFonts.ReceiveNotificationsOfFonts.pdf");
            //ExEnd:ReceiveNotificationsOfFonts
        }

        [Test]
        public void ReceiveWarningNotification()
        {
            //ExStart:ReceiveWarningNotification
            Document doc = new Document(MyDir + "Rendering.docx");
            
            // When you call UpdatePageLayout the document is rendered in memory. Any warnings that occured during rendering
            // are stored until the document save and then sent to the appropriate WarningCallback.
            doc.UpdatePageLayout();

            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.WarningCallback = callback;
            
            // Even though the document was rendered previously, any save warnings are notified to the user during document save.
            doc.Save(ArtifactsDir + "WorkingWithFonts.ReceiveWarningNotification.pdf");
            //ExEnd:ReceiveWarningNotification  
        }

        //ExStart:HandleDocumentWarnings
        public class HandleDocumentWarnings : IWarningCallback
        {
            /// <summary>
            /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
            /// Potential issue during document procssing. The callback can be set to listen for warnings generated
            /// during document load and/or document save.
            /// </summary>
            public void Warning(WarningInfo info)
            {
                // We are only interested in fonts being substituted.
                if (info.WarningType == WarningType.FontSubstitution)
                {
                    Console.WriteLine("Font substitution: " + info.Description);
                }
            }
        }
        //ExEnd:HandleDocumentWarnings

        [Test]
        //ExStart:ResourceSteam
        //GistId:7e64f6d40825be58a8c12f1307c12964
        public void ResourceSteam()
        {
            Document doc = new Document(MyDir + "Rendering.docx");
            
            FontSettings.DefaultInstance.SetFontsSources(new FontSourceBase[]
                { new SystemFontSource(), new ResourceSteamFontSource() });

            doc.Save(ArtifactsDir + "WorkingWithFonts.ResourceSteam.pdf");
        }

        internal class ResourceSteamFontSource : StreamFontSource
        {
            public override Stream OpenFontDataStream()
            {
                return Assembly.GetExecutingAssembly().GetManifestResourceStream("resourceName");
            }
        }
        //ExEnd:ResourceSteam

        [Test]
        //ExStart:GetSubstitutionWithoutSuffixes
        //GistId:a08698f540d47082b4e2dbb1cb67fc1b
        public void GetSubstitutionWithoutSuffixes()
        {
            Document doc = new Document(MyDir + "Get substitution without suffixes.docx");

            DocumentSubstitutionWarnings substitutionWarningHandler = new DocumentSubstitutionWarnings();
            doc.WarningCallback = substitutionWarningHandler;

            List<FontSourceBase> fontSources = new List<FontSourceBase>(FontSettings.DefaultInstance.GetFontsSources());

            FolderFontSource folderFontSource = new FolderFontSource(FontsDir, true);
            fontSources.Add(folderFontSource);

            FontSourceBase[] updatedFontSources = fontSources.ToArray();
            FontSettings.DefaultInstance.SetFontsSources(updatedFontSources);

            doc.Save(ArtifactsDir + "WorkingWithFonts.GetSubstitutionWithoutSuffixes.pdf");

            Assert.That(substitutionWarningHandler.FontWarnings[0].Description, Is.EqualTo("Font 'DINOT-Regular' has not been found. Using 'DINOT' font instead. Reason: font name substitution."));
        }

        public class DocumentSubstitutionWarnings : IWarningCallback
        {
            /// <summary>
            /// Our callback only needs to implement the "Warning" method.
            /// This method is called whenever there is a potential issue during document processing.
            /// The callback can be set to listen for warnings generated during document load and/or document save.
            /// </summary>
            public void Warning(WarningInfo info)
            {
                // We are only interested in fonts being substituted.
                if (info.WarningType == WarningType.FontSubstitution)
                    FontWarnings.Warning(info);
            }

            public WarningInfoCollection FontWarnings = new WarningInfoCollection();
        }
        //ExEnd:GetSubstitutionWithoutSuffixes
    }
}