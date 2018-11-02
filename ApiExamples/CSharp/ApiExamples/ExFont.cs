// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Fonts;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExFont : ApiExampleBase
    {
        [Test]
        public void CreateFormattedRun()
        {
            //ExStart
            //ExFor:Document.#ctor
            //ExFor:Font
            //ExFor:Font.Name
            //ExFor:Font.Size
            //ExFor:Font.HighlightColor
            //ExFor:Run
            //ExFor:Run.#ctor(DocumentBase,String)
            //ExFor:Story.FirstParagraph
            //ExSummary:Shows how to add a formatted run of text to a document using the object model.
            // Create an empty document. It contains one empty paragraph.
            Document doc = new Document();

            // Create a new run of text.
            Run run = new Run(doc, "Hello");

            // Specify character formatting for the run of text.
            Aspose.Words.Font f = run.Font;
            f.Name = "Courier New";
            f.Size = 36;
            f.HighlightColor = Color.Yellow;

            // Append the run of text to the end of the first paragraph
            // in the body of the first section of the document.
            doc.FirstSection.Body.FirstParagraph.AppendChild(run);
            //ExEnd
        }

        [Test]
        public void Caps()
        {
            //ExStart
            //ExFor:Font.AllCaps
            //ExFor:Font.SmallCaps
            //ExSummary:Shows how to use all capitals and small capitals character formatting properties.
            // Create an empty document. It contains one empty paragraph.
            Document doc = new Document();

            // Get the paragraph from the document, we will be adding runs of text to it.
            Paragraph para = (Paragraph) doc.GetChild(NodeType.Paragraph, 0, true);

            Run run = new Run(doc, "All capitals");
            run.Font.AllCaps = true;
            para.AppendChild(run);

            run = new Run(doc, "SMALL CAPITALS");
            run.Font.SmallCaps = true;
            para.AppendChild(run);
            //ExEnd
        }

        [Test]
        public void GetDocumentFonts()
        {
            //ExStart
            //ExFor:FontInfoCollection
            //ExFor:DocumentBase.FontInfos
            //ExFor:FontInfo
            //ExFor:FontInfo.Name
            //ExFor:FontInfo.IsTrueType
            //ExSummary:Shows how to gather the details of what fonts are present in a document.
            Document doc = new Document(MyDir + "Document.doc");

            FontInfoCollection fonts = doc.FontInfos;
            int fontIndex = 1;

            // The fonts info extracted from this document does not necessarily mean that the fonts themselves are
            // used in the document. If a font is present but not used then most likely they were referenced at some time
            // and then removed from the Document.
            foreach (FontInfo info in fonts)
            {
                // Print out some important details about the font.
                Console.WriteLine("Font #{0}", fontIndex);
                Console.WriteLine("Name: {0}", info.Name);
                Console.WriteLine("IsTrueType: {0}", info.IsTrueType);
                fontIndex++;
            }

            //ExEnd
        }

        [Test]
        [Description("WORDSNET-16234")]
        public void DefaultValuesEmbeddedFontsParameters()
        {
            Document doc = new Document();

            Assert.IsFalse(doc.FontInfos.EmbedTrueTypeFonts);
            Assert.IsFalse(doc.FontInfos.EmbedSystemFonts);
            Assert.IsFalse(doc.FontInfos.SaveSubsetFonts);
        }

        [Test]
        public void FontInfoCollection()
        {
            //ExStart
            //ExFor:FontInfoCollection
            //ExFor:DocumentBase.FontInfos
            //ExFor:FontInfoCollection.EmbedTrueTypeFonts
            //ExFor:FontInfoCollection.EmbedSystemFonts
            //ExFor:FontInfoCollection.SaveSubsetFonts
            //ExSummary:Shows how to save a document with embedded TrueType fonts
            Document doc = new Document(MyDir + "Document.docx");

            FontInfoCollection fontInfos = doc.FontInfos;
            fontInfos.EmbedTrueTypeFonts = true;
            fontInfos.EmbedSystemFonts = false;
            fontInfos.SaveSubsetFonts = false;

            doc.Save(MyDir + @"/Artifacts/Document.docx");
            //ExEnd
        }

        [Test]
        [TestCase(true, false, false, Description =
            "Save a document with embedded TrueType fonts. System fonts are not included. Saves full versions of embedding fonts.")]
        [TestCase(true, true, false, Description =
            "Save a document with embedded TrueType fonts. System fonts are included. Saves full versions of embedding fonts.")]
        [TestCase(true, true, true, Description =
            "Save a document with embedded TrueType fonts. System fonts are included. Saves subset of embedding fonts.")]
        [TestCase(true, false, true, Description =
            "Save a document with embedded TrueType fonts. System fonts are not included. Saves subset of embedding fonts.")]
        [TestCase(false, false, false, Description = "Remove embedded fonts from the saved document.")]
        public void WorkWithEmbeddedFonts(bool embedTrueTypeFonts, bool embedSystemFonts, bool saveSubsetFonts)
        {
            Document doc = new Document(MyDir + "Document.doc");

            FontInfoCollection fontInfos = doc.FontInfos;
            fontInfos.EmbedTrueTypeFonts = embedTrueTypeFonts;
            fontInfos.EmbedSystemFonts = embedSystemFonts;
            fontInfos.SaveSubsetFonts = saveSubsetFonts;

            doc.Save(MyDir + @"/Artifacts/Document.docx");
        }

        [Test]
        public void StrikeThrough()
        {
            //ExStart
            //ExFor:Font.StrikeThrough
            //ExFor:Font.DoubleStrikeThrough
            //ExSummary:Shows how to use strike-through character formatting properties.
            // Create an empty document. It contains one empty paragraph.
            Document doc = new Document();

            // Get the paragraph from the document, we will be adding runs of text to it.
            Paragraph para = (Paragraph) doc.GetChild(NodeType.Paragraph, 0, true);

            Run run = new Run(doc, "Double strike through text");
            run.Font.DoubleStrikeThrough = true;
            para.AppendChild(run);

            run = new Run(doc, "Single strike through text");
            run.Font.StrikeThrough = true;
            para.AppendChild(run);
            //ExEnd
        }

        [Test]
        public void PositionSubscript()
        {
            //ExStart
            //ExFor:Font.Position
            //ExFor:Font.Subscript
            //ExFor:Font.Superscript
            //ExSummary:Shows how to use subscript, superscript, complex script, text effects, and baseline text position properties.
            // Create an empty document. It contains one empty paragraph.
            Document doc = new Document();

            // Get the paragraph from the document, we will be adding runs of text to it.
            Paragraph para = (Paragraph) doc.GetChild(NodeType.Paragraph, 0, true);

            // Add a run of text that is raised 5 points above the baseline.
            Run run = new Run(doc, "Raised text");
            run.Font.Position = 5;
            para.AppendChild(run);

            // Add a run of normal text.
            run = new Run(doc, "Normal text");
            para.AppendChild(run);

            // Add a run of text that appears as subscript.
            run = new Run(doc, "Subscript");
            run.Font.Subscript = true;
            para.AppendChild(run);

            // Add a run of text that appears as superscript.
            run = new Run(doc, "Superscript");
            run.Font.Superscript = true;
            para.AppendChild(run);
            //ExEnd
        }

        [Test]
        public void ScalingSpacing()
        {
            //ExStart
            //ExFor:Font.Scaling
            //ExFor:Font.Spacing
            //ExSummary:Shows how to use character scaling and spacing properties.
            // Create an empty document. It contains one empty paragraph.
            Document doc = new Document();

            // Get the paragraph from the document, we will be adding runs of text to it.
            Paragraph para = (Paragraph) doc.GetChild(NodeType.Paragraph, 0, true);

            // Add a run of text with characters 150% width of normal characters.
            Run run = new Run(doc, "Wide characters");
            run.Font.Scaling = 150;
            para.AppendChild(run);

            // Add a run of text with extra 1pt space between characters.
            run = new Run(doc, "Expanded by 1pt");
            run.Font.Spacing = 1;
            para.AppendChild(run);

            // Add a run of text with space between characters reduced by 1pt.
            run = new Run(doc, "Condensed by 1pt");
            run.Font.Spacing = -1;
            para.AppendChild(run);
            //ExEnd
        }

        [Test]
        public void EmbossItalic()
        {
            Document doc = new Document();
            //ExStart
            //ExFor:Font.Emboss
            //ExFor:Font.Italic
            //ExSummary:Shows how to create a run of formatted text.
            Run run = new Run(doc, "Hello");
            run.Font.Emboss = true;
            run.Font.Italic = true;
            //ExEnd
        }

        [Test]
        public void Engrave()
        {
            Document doc = new Document();
            //ExStart
            //ExFor:Font.Engrave
            //ExSummary:Shows how to create a run of text formatted as engraved.
            Run run = new Run(doc, "Hello");
            run.Font.Engrave = true;
            //ExEnd
        }

        [Test]
        public void Shadow()
        {
            Document doc = new Document();
            //ExStart
            //ExFor:Font.Shadow
            //ExSummary:Shows how to create a run of text formatted with a shadow.
            Run run = new Run(doc, "Hello");
            run.Font.Shadow = true;
            //ExEnd
        }

        [Test]
        public void Outline()
        {
            Document doc = new Document();
            //ExStart
            //ExFor:Font.Outline
            //ExSummary:Shows how to create a run of text formatted as outline.
            Run run = new Run(doc, "Hello");
            run.Font.Outline = true;
            //ExEnd
        }

        [Test]
        public void Hidden()
        {
            Document doc = new Document();
            //ExStart
            //ExFor:Font.Hidden
            //ExSummary:Shows how to create a hidden run of text.
            Run run = new Run(doc, "Hello");
            run.Font.Hidden = true;
            //ExEnd
        }

        [Test]
        public void Kerning()
        {
            Document doc = new Document();
            //ExStart
            //ExFor:Font.Kerning
            //ExSummary:Shows how to specify the font size at which kerning starts.
            Run run = new Run(doc, "Hello");
            run.Font.Kerning = 24;
            //ExEnd
        }

        [Test]
        public void NoProofing()
        {
            Document doc = new Document();
            //ExStart
            //ExFor:Font.NoProofing
            //ExSummary:Shows how to specify that the run of text is not to be spell checked by Microsoft Word.
            Run run = new Run(doc, "Hello");
            run.Font.NoProofing = true;
            //ExEnd
        }

        [Test]
        public void LocaleId()
        {
            Document doc = new Document();

            //ExStart
            //ExFor:Font.LocaleId
            //ExSummary:Shows how to specify the language of a text run so Microsoft Word can use a proper spell checker.
            //Create a run of text that contains Russian text.
            Run run = new Run(doc, "Привет");

            //Specify the locale so Microsoft Word recognizes this text as Russian.
            //For the list of locale identifiers see https://docs.microsoft.com/en-us/deployoffice/office2016/language-identifiers-and-optionstate-id-values-in-office-2016
            run.Font.LocaleId = 1049;
            //ExEnd
        }

        [Test]
        public void Underlines()
        {
            Document doc = new Document();
            //ExStart
            //ExFor:Font.Underline
            //ExFor:Font.UnderlineColor
            //ExSummary:Shows how use the underline character formatting properties.
            Run run = new Run(doc, "Hello");
            run.Font.Underline = Underline.Dotted;
            run.Font.UnderlineColor = Color.Red;
            //ExEnd
        }

        [Test]
        public void ComplexScript()
        {
            Document doc = new Document();
            //ExStart
            //ExFor:Font.ComplexScript
            //ExSummary:Shows how to make a run that's always treated as complex script.
            Run run = new Run(doc, "Complex script");
            run.Font.ComplexScript = true;
            //ExEnd
        }

        [Test]
        public void SparklingText()
        {
            Document doc = new Document();
            //ExStart
            //ExFor:Font.TextEffect
            //ExSummary:Shows how to apply a visual effect to a run.
            Run run = new Run(doc, "Text with effect");
            run.Font.TextEffect = TextEffect.SparkleText;
            //ExEnd
        }

        [Test]
        public void Shading()
        {
            //ExStart
            //ExFor:Font.Shading
            //ExSummary:Shows how to apply shading for a run of text.
            DocumentBuilder builder = new DocumentBuilder();

            Shading shd = builder.Font.Shading;
            shd.Texture = TextureIndex.TextureDiagonalCross;
            shd.BackgroundPatternColor = Color.Blue;
            shd.ForegroundPatternColor = Color.BlueViolet;

            builder.Font.Color = Color.White;

            builder.Writeln("White text on a blue background with texture.");
            //ExEnd
        }

        [Test]
        public void Bidi()
        {
            //ExStart
            //ExFor:Font.Bidi
            //ExFor:Font.NameBi
            //ExFor:Font.SizeBi
            //ExFor:Font.ItalicBi
            //ExFor:Font.BoldBi
            //ExFor:Font.LocaleIdBi
            //ExSummary:Shows how to insert and format right-to-left text.
            DocumentBuilder builder = new DocumentBuilder();

            // Signal to Microsoft Word that this run of text contains right-to-left text.
            builder.Font.Bidi = true;

            // Specify the font and font size to be used for the right-to-left text.
            builder.Font.NameBi = "Andalus";
            builder.Font.SizeBi = 48;

            // Specify that the right-to-left text in this run is bold and italic.
            builder.Font.ItalicBi = true;
            builder.Font.BoldBi = true;

            // Specify the locale so Microsoft Word recognizes this text as Arabic - Saudi Arabia.
            // For the list of locale identifiers see https://docs.microsoft.com/en-us/deployoffice/office2016/language-identifiers-and-optionstate-id-values-in-office-2016
            builder.Font.LocaleIdBi = 1025;

            // Insert some Arabic text.
            builder.Writeln("مرحبًا");

            builder.Document.Save(MyDir + @"\Artifacts\Font.Bidi.doc");
            //ExEnd
        }

        [Test]
        public void FarEast()
        {
            //ExStart
            //ExFor:Font.NameFarEast
            //ExFor:Font.LocaleIdFarEast
            //ExSummary:Shows how to insert and format text in Chinese or any other Far East language.
            DocumentBuilder builder = new DocumentBuilder();

            builder.Font.Size = 48;

            // Specify the font name. Make sure it the font has the glyphs that you want to display.
            builder.Font.NameFarEast = "SimSun";

            // Specify the locale so Microsoft Word recognizes this text as Chinese.
            // For the list of locale identifiers see https://docs.microsoft.com/en-us/deployoffice/office2016/language-identifiers-and-optionstate-id-values-in-office-2016
            builder.Font.LocaleIdFarEast = 2052;

            // Insert some Chinese text.
            builder.Writeln("你好世界");

            builder.Document.Save(MyDir + @"\Artifacts\Font.FarEast.doc");
            //ExEnd
        }

        [Test]
        public void Names()
        {
            //ExStart
            //ExFor:Font.NameAscii
            //ExFor:Font.NameOther
            //ExSummary:A pretty unusual example of how Microsoft Word can combine two different fonts in one run.
            DocumentBuilder builder = new DocumentBuilder();

            // This tells Microsoft Word to use Arial for characters 0..127 and
            // Times New Roman for characters 128..255. 
            // Looks like a pretty strange case to me, but it is possible.
            builder.Font.NameAscii = "Arial";
            builder.Font.NameOther = "Times New Roman";

            builder.Writeln("Hello, Привет");

            builder.Document.Save(MyDir + @"\Artifacts\Font.Names.doc");
            //ExEnd
        }

        [Test]
        public void ChangeStyleIdentifier()
        {
            //ExStart
            //ExFor:Font.StyleIdentifier
            //ExFor:StyleIdentifier
            //ExSummary:Shows how to use style identifier to find text formatted with a specific character style and apply different character style.
            Document doc = new Document(MyDir + "Font.StyleIdentifier.doc");

            // Select all run nodes in the document.
            NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);

            // Loop through every run node.
            foreach (Run run in runs.OfType<Run>())
            {
                // If the character style of the run is what we want, do what we need. Change the style in this case.
                // Note that using StyleIdentifier we can identify a built-in style regardless 
                // of the language of Microsoft Word used to create the document.
                if (run.Font.StyleIdentifier.Equals(StyleIdentifier.Emphasis))
                    run.Font.StyleIdentifier = StyleIdentifier.Strong;
            }

            doc.Save(MyDir + @"\Artifacts\Font.StyleIdentifier.doc");
            //ExEnd
        }

        [Test]
        public void ChangeStyleName()
        {
            //ExStart
            //ExFor:Font.StyleName
            //ExSummary:Shows how to use style name to find text formatted with a specific character style and apply different character style.
            Document doc = new Document(MyDir + "Font.StyleName.doc");

            // Select all run nodes in the document.
            NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);

            // Loop through every run node.
            foreach (Run run in runs.OfType<Run>())
            {
                // If the character style of the run is what we want, do what we need. Change the style in this case.
                // Note that names of built in styles could be different in documents 
                // created by Microsoft Word versions for different languages.
                if (run.Font.StyleName.Equals("Emphasis"))
                    run.Font.StyleName = "Strong";
            }

            doc.Save(MyDir + @"\Artifacts\Font.StyleName.doc");
            //ExEnd
        }

        [Test]
        public void Style()
        {
            //ExStart
            //ExFor:Font.Style
            //ExFor:Style.BuiltIn
            //ExSummary:Applies double underline to all runs in a document that are formatted with custom character styles.
            Document doc = new Document(MyDir + "Font.Style.doc");

            // Select all run nodes in the document.
            NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);

            // Loop through every run node.
            foreach (Run run in runs.OfType<Run>())
            {
                Style charStyle = run.Font.Style;

                // If the style of the run is not a built-in character style, apply double underline.
                if (!charStyle.BuiltIn)
                    run.Font.Underline = Underline.Double;
            }

            doc.Save(MyDir + @"\Artifacts\Font.Style.doc");
            //ExEnd
        }

        [Test]
        public void GetAllFonts()
        {
            //ExStart
            //ExFor:Run
            //ExSummary:Gets all fonts used in a document.
            Document doc = new Document(MyDir + "Font.Names.doc");

            // Select all runs in the document.
            NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);

            // Use a hashtable so we will keep only unique font names.
            Hashtable fontNames = new Hashtable();

            foreach (Run run in runs.OfType<Run>())
            {
                // This adds an entry into the hashtable.
                // The key is the font name. The value is null, we don't need the value.
                fontNames[run.Font.Name] = null;
            }

            // There are two fonts used in this document.
            Console.WriteLine("Font Count: " + fontNames.Count);
            //ExEnd

            // Verify the font count is correct.
            Assert.AreEqual(2, fontNames.Count);
        }

        [Test]
        public void RecieveFontSubstitutionNotification()
        {
            // Store the font sources currently used so we can restore them later. 
            FontSourceBase[] origFontSources = FontSettings.DefaultInstance.GetFontsSources();

            //ExStart
            //ExFor:IWarningCallback
            //ExFor:DocumentBase.WarningCallback
            //ExFor:SaveOptions.WarningCallback
            //ExFor:Fonts.FontSettings.DefaultInstance
            //ExId:FontSubstitutionNotification
            //ExSummary:Demonstrates how to receive notifications of font substitutions by using IWarningCallback.
            // Load the document to render.
            Document doc = new Document(MyDir + "Document.doc");

            // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class.
            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.WarningCallback = callback;

            // We can choose the default font to use in the case of any missing fonts.
            FontSettings.DefaultInstance.DefaultFontName = "Arial";

            // For testing we will set Aspose.Words to look for fonts only in a folder which doesn't exist. Since Aspose.Words won't
            // find any fonts in the specified directory, then during rendering the fonts in the document will be substituted with the default 
            // font specified under FontSettings.DefaultFontName. We can pick up on this substitution using our callback.
            FontSettings.DefaultInstance.SetFontsFolder(String.Empty, false);

            // Pass the save options along with the save path to the save method.
            doc.Save(MyDir + @"\Artifacts\Rendering.MissingFontNotification.pdf");
            //ExEnd

            Assert.Greater(callback.mFontWarnings.Count, 0);
            Assert.True(callback.mFontWarnings[0].WarningType == WarningType.FontSubstitution);
            Assert.True(callback.mFontWarnings[0].Description
                .Equals(
                    "Font 'Times New Roman' has not been found. Using 'Fanwood' font instead. Reason: first available font."));

            // Restore default fonts. 
            FontSettings.DefaultInstance.SetFontsSources(origFontSources);
        }

        [Test]
        public void GetAvailableFonts()
        {
            //ExStart
            //ExFor:FontSourceBase.GetAvailableFonts
            //ExFor:PhysicalFontInfo.FontFamilyName
            //ExFor:PhysicalFontInfo.FullFontName
            //ExFor:PhysicalFontInfo.Version
            //ExFor:PhysicalFontInfo.FilePath
            //ExSummary:Shows how to get available fonts and information about them.
            // Add a new folder source which will instruct Aspose.Words to search the following folder for fonts. 
            FontSourceBase[] folderFontSource = { new FolderFontSource(MyDir + @"MyFonts\", true) };
            
            foreach (PhysicalFontInfo fontInfo in folderFontSource[0].GetAvailableFonts())
            {
                Console.WriteLine("FontFamilyName : {0}", fontInfo.FontFamilyName);
                Console.WriteLine("FullFontName  : {0}", fontInfo.FullFontName);
                Console.WriteLine("Version  : {0}", fontInfo.Version);
                Console.WriteLine("FilePath : {0}\n", fontInfo.FilePath);
            }
            //ExEnd
        }

        //ExStart
        //ExFor:IWarningCallback
        //ExFor:DocumentBase.WarningCallback
        //ExFor:SaveOptions.WarningCallback
        //ExId:FontSubstitutionWarningCallback
        //ExSummary:Shows how to implement the IWarningCallback to be notified of any font substitution during document save.
        public class HandleDocumentWarnings : IWarningCallback
        {
            /// <summary>
            /// Our callback only needs to implement the "Warning" method. This method is called whenever there is a
            /// potential issue during document processing. The callback can be set to listen for warnings generated during document
            /// load and/or document save.
            /// </summary>
            public void Warning(WarningInfo info)
            {
                // We are only interested in fonts being substituted.
                if (info.WarningType == WarningType.FontSubstitution)
                {
                    Console.WriteLine("Font substitution: " + info.Description);
                    mFontWarnings.Warning(info); //ExSkip
                }
            }

            public WarningInfoCollection mFontWarnings = new WarningInfoCollection(); //ExSkip
        }
        //ExEnd

        [Test]
        public void EnableFontSubstitutionTrue()
        {
            //ExStart
            //ExFor:FontSettings.EnableFontSubstitution
            //ExSummary:Shows how to set the property for finding the closest match font among the available font sources instead missing font.
            Document doc = new Document(MyDir + "Font.EnableFontSubstitution.docx");

            // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class.
            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.WarningCallback = callback;

            FontSettings fontSettings = new FontSettings();
            fontSettings.DefaultFontName = "Arial";
            fontSettings.EnableFontSubstitution = true;
            //ExEnd

            doc.FontSettings = fontSettings;

            doc.Save(MyDir + @"\Artifacts\Font.EnableFontSubstitution.pdf");

            Assert.True(callback.mFontWarnings[0].Description
                .Equals(
                    "Font '28 Days Later' has not been found. Using 'Franklin Gothic Medium' font instead. Reason: closest match according to font info from the document."));
        }

        [Test]
        public void EnableFontSubstitutionFalse()
        {
            Document doc = new Document(MyDir + "Font.EnableFontSubstitution.docx");

            // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class.
            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.WarningCallback = callback;

            FontSettings fontSettings = new FontSettings();
            fontSettings.DefaultFontName = "Arial";
            fontSettings.EnableFontSubstitution = false;

            doc.FontSettings = fontSettings;

            doc.Save(MyDir + @"\Artifacts\Font.EnableFontSubstitution.pdf");

            Assert.True(callback.mFontWarnings[0].Description
                .Equals(
                    "Font '28 Days Later' has not been found. Using 'Arial' font instead. Reason: default font setting."));
        }

        [Test]
        public void FontSubstitutionWarnings()
        {
            Document doc = new Document(MyDir + "Rendering.doc");

            // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class.
            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.WarningCallback = callback;

            FontSettings fontSettings = new FontSettings();
            fontSettings.DefaultFontName = "Arial";
            fontSettings.SetFontSubstitutes("Arial", new string[] { "Arvo", "Slab" });
            fontSettings.SetFontsFolder(MyDir + @"MyFonts\", false);

            doc.FontSettings = fontSettings;

            doc.Save(MyDir + @"\Artifacts\Rendering.MissingFontNotification.pdf");

            Assert.True(callback.mFontWarnings[0].Description
                .Equals("Font substitutes: 'Arial' replaced with 'Arvo'."));
            Assert.True(callback.mFontWarnings[1].Description
                .Equals(
                    "Font 'Times New Roman' has not been found. Using 'Noticia Text' font instead. Reason: closest match according to font info from the document."));
        }

        [Test]
        public void FontSubstitutionWarningsClosestMatch()
        {
            Document doc = new Document(MyDir + "Font.DisappearingBulletPoints.doc");

            // Create a new class implementing IWarningCallback and assign it to the PdfSaveOptions class.
            HandleDocumentWarnings callback = new HandleDocumentWarnings();
            doc.WarningCallback = callback;

            doc.Save(MyDir + @"\Artifacts\Font.DisappearingBulletPoints.pdf");

            Assert.True(callback.mFontWarnings[0].Description
                .Equals(
                    "Font 'SymbolPS' has not been found. Using 'Wingdings' font instead. Reason: closest match according to font info from the document."));
        }

        [Test]
        public void SetFontAutoColor()
        {
            //ExStart
            //ExFor:Font.AutoColor
            //ExSummary:Shows how calculated color of the text (black or white) to be used for 'auto color'
            Run run = new Run(new Document());

            // Remove direct color, so it can be calculated automatically with Font.AutoColor.
            run.Font.Color = Color.Empty;

            // When we set black color for background, autocolor for font must be white
            run.Font.Shading.BackgroundPatternColor = Color.Black;
            Assert.AreEqual(Color.White, run.Font.AutoColor);

            // When we set white color for background, autocolor for font must be black
            run.Font.Shading.BackgroundPatternColor = Color.White;
            Assert.AreEqual(Color.Black, run.Font.AutoColor);
            //ExEnd
        }

        [Test]
        public void RemoveHiddenContentFromDocument()
        {
            //ExStart
            //ExFor:Font.Hidden
            //ExFor:Paragraph.Accept
            //ExFor:DocumentVisitor.VisitParagraphStart(Paragraph)
            //ExFor:DocumentVisitor.VisitFormField(FormField)
            //ExFor:DocumentVisitor.VisitTableEnd(Table)
            //ExFor:DocumentVisitor.VisitCellEnd(Cell)
            //ExFor:DocumentVisitor.VisitRowEnd(Row)
            //ExFor:DocumentVisitor.VisitSpecialChar(SpecialChar)
            //ExFor:DocumentVisitor.VisitGroupShapeStart(GroupShape)
            //ExFor:DocumentVisitor.VisitShapeStart(Shape)
            //ExFor:DocumentVisitor.VisitCommentStart(Comment)
            //ExFor:DocumentVisitor.VisitFootnoteStart(Footnote)
            //ExFor:SpecialChar
            //ExFor:Node.Accept
            //ExFor:Paragraph.ParagraphBreakFont
            //ExFor:Table.Accept
            //ExSummary:Implements the Visitor Pattern to remove all content formatted as hidden from the document.
            // Open the document we want to remove hidden content from.
            Document doc = new Document(MyDir + "Font.Hidden.doc");

            // Create an object that inherits from the DocumentVisitor class.
            RemoveHiddenContentVisitor hiddenContentRemover = new RemoveHiddenContentVisitor();

            // This is the well known Visitor pattern. Get the model to accept a visitor.
            // The model will iterate through itself by calling the corresponding methods
            // on the visitor object (this is called visiting).

            // We can run it over the entire the document like so:
            doc.Accept(hiddenContentRemover);

            // Or we can run it on only a specific node.
            Paragraph para = (Paragraph) doc.GetChild(NodeType.Paragraph, 4, true);
            para.Accept(hiddenContentRemover);

            // Or over a different type of node like below.
            Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
            table.Accept(hiddenContentRemover);

            doc.Save(MyDir + @"\Artifacts\Font.Hidden.doc");

            Assert.AreEqual(13, doc.GetChildNodes(NodeType.Paragraph, true).Count); //ExSkip
            Assert.AreEqual(1, doc.GetChildNodes(NodeType.Table, true).Count); //ExSkip
        }

        /// <summary>
        /// This class when executed will remove all hidden content from the Document. Implemented as a Visitor.
        /// </summary>
        class RemoveHiddenContentVisitor : DocumentVisitor
        {
            /// <summary>
            /// Called when a FieldStart node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldStart(FieldStart fieldStart)
            {
                // If this node is hidden, then remove it.
                if (isHidden(fieldStart))
                    fieldStart.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldEnd node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldEnd(FieldEnd fieldEnd)
            {
                if (isHidden(fieldEnd))
                    fieldEnd.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldSeparator node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldSeparator(FieldSeparator fieldSeparator)
            {
                if (isHidden(fieldSeparator))
                    fieldSeparator.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Run node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitRun(Run run)
            {
                if (isHidden(run))
                    run.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Paragraph node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitParagraphStart(Paragraph paragraph)
            {
                if (isHidden(paragraph))
                    paragraph.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FormField is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFormField(FormField formField)
            {
                if (isHidden(formField))
                    formField.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a GroupShape is encountered in the document.
            /// </summary>
            public override VisitorAction VisitGroupShapeStart(GroupShape groupShape)
            {
                if (isHidden(groupShape))
                    groupShape.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Shape is encountered in the document.
            /// </summary>
            public override VisitorAction VisitShapeStart(Shape shape)
            {
                if (isHidden(shape))
                    shape.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Comment is encountered in the document.
            /// </summary>
            public override VisitorAction VisitCommentStart(Comment comment)
            {
                if (isHidden(comment))
                    comment.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Footnote is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFootnoteStart(Footnote footnote)
            {
                if (isHidden(footnote))
                    footnote.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when visiting of a Table node is ended in the document.
            /// </summary>
            public override VisitorAction VisitTableEnd(Table table)
            {
                // At the moment there is no way to tell if a particular Table/Row/Cell is hidden. 
                // Instead, if the content of a table is hidden, then all inline child nodes of the table should be 
                // hidden and thus removed by previous visits as well. This will result in the container being empty
                // so if this is the case we know to remove the table node.
                //
                // Note that a table which is not hidden but simply has no content will not be affected by this algorithm,
                // as technically they are not completely empty (for example a properly formed Cell will have at least 
                // an empty paragraph in it)
                if (!table.HasChildNodes)
                    table.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when visiting of a Cell node is ended in the document.
            /// </summary>
            public override VisitorAction VisitCellEnd(Cell cell)
            {
                if (!cell.HasChildNodes && cell.ParentNode != null)
                    cell.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when visiting of a Row node is ended in the document.
            /// </summary>
            public override VisitorAction VisitRowEnd(Row row)
            {
                if (!row.HasChildNodes && row.ParentNode != null)
                    row.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a SpecialCharacter is encountered in the document.
            /// </summary>
            public override VisitorAction VisitSpecialChar(SpecialChar specialChar)
            {
                if (isHidden(specialChar))
                    specialChar.Remove();

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Returns true if the node passed is set as hidden, returns false if it is visible.
            /// </summary>
            private bool isHidden(Node node)
            {
                if (node is Inline currentNode)
                {
                    // If the node is Inline then cast it to retrieve the Font property which contains the hidden property
                    return currentNode.Font.Hidden;
                }
                else if (node.NodeType == NodeType.Paragraph)
                {
                    // If the node is a paragraph cast it to retrieve the ParagraphBreakFont which contains the hidden property
                    Paragraph para = (Paragraph) node;
                    return para.ParagraphBreakFont.Hidden;
                }
                else if (node is ShapeBase shape)
                {
                    // Node is a shape or groupshape.
                    return shape.Font.Hidden;
                }
                else if (node is InlineStory inlineStory)
                {
                    // Node is a comment or footnote.
                    return inlineStory.Font.Hidden;
                }

                // A node that is passed to this method which does not contain a hidden property will end up here. 
                // By default nodes are not hidden so return false.
                return false;
            }
        }

        //ExEnd

        [Test]
        public void BlankDocumentFonts()
        {
            //ExStart
            //ExFor:Fonts.FontInfoCollection.Contains(String)
            //ExFor:Fonts.FontInfoCollection.Count
            //ExSummary:Shows info about the fonts that are present in the blank document.
            // Create a new document
            Document doc = new Document();
            // A blank document comes with 3 fonts
            Assert.AreEqual(3, doc.FontInfos.Count);
            Assert.AreEqual(true, doc.FontInfos.Contains("Times New Roman"));
            Assert.AreEqual(true, doc.FontInfos.Contains("Symbol"));
            Assert.AreEqual(true, doc.FontInfos.Contains("Arial"));
            //ExEnd
        }

        [Test]
        public void ExtractEmbeddedFont()
        {
            //ExStart
            //ExFor:Fonts.EmbeddedFontFormat
            //ExFor:Fonts.EmbeddedFontStyle
            //ExFor:Fonts.FontInfo.GetEmbeddedFont(EmbeddedFontFormat,EmbeddedFontStyle)
            //ExFor:Fonts.FontInfoCollection.Item(Int32)
            //ExFor:Fonts.FontInfoCollection.Item(String)
            //ExSummary:Shows how to extract embedded font from a document.
            Document doc = new Document(MyDir + "Font.Embedded.docx");
            // Let's get the font we are interested in
            FontInfo mittelschriftInfo = doc.FontInfos[2];
            // We can now extract this embedded font
            byte[] embeddedFontBytes = mittelschriftInfo.GetEmbeddedFont(EmbeddedFontFormat.OpenType, EmbeddedFontStyle.Regular);
            Assert.IsNotNull(embeddedFontBytes);
            // Then we can save the font to our directory
            File.WriteAllBytes(MyDir + @"\Artifacts\Alte DIN 1451 Mittelschrift.ttf", embeddedFontBytes);
            
            // If we want to extract a font from a .doc as opposed to a .docx, we need to make sure to set the appropriate embedded font format
            doc = new Document(MyDir + "Font.Embedded.doc");

            Assert.IsNull(doc.FontInfos["Alte DIN 1451 Mittelschrift"].GetEmbeddedFont(EmbeddedFontFormat.OpenType, EmbeddedFontStyle.Regular));
            Assert.IsNotNull(doc.FontInfos["Alte DIN 1451 Mittelschrift"].GetEmbeddedFont(EmbeddedFontFormat.EmbeddedOpenType, EmbeddedFontStyle.Regular));
            //ExEnd
        }

        [Test]
        public void GetFontInfoFromFile() 
        {
            //ExStart
            //ExFor:Fonts.FontFamily
            //ExFor:Fonts.FontPitch
            //ExFor:Fonts.FontInfo.AltName
            //ExFor:Fonts.FontInfo.Charset
            //ExFor:Fonts.FontInfo.Family
            //ExFor:Fonts.FontInfo.Panose
            //ExFor:Fonts.FontInfo.Pitch
            //ExFor:Fonts.FontInfoCollection.GetEnumerator
            //ExSummary:Shows how to get information about each font in a document.
            Document doc = new Document(MyDir + "Font.Embedded.docx");
            
            // We can iterate over all the fonts with an enumerator
            IEnumerator fontCollectionEnumerator = doc.FontInfos.GetEnumerator();
            // Print detailed information about each font to the console
            while (fontCollectionEnumerator.MoveNext())
            {
                FontInfo fontInfo = (FontInfo)fontCollectionEnumerator.Current;
                if (fontInfo != null)
                {
                    Console.WriteLine("Font name: " + fontInfo.Name);
                    Console.WriteLine("Alt name: " + fontInfo.AltName); // Alt names are usually blank
                    Console.WriteLine("\t- Family: " + fontInfo.Family);
                    Console.WriteLine("\t- " + (fontInfo.IsTrueType ? "Is TrueType" : "Is not TrueType"));
                    Console.WriteLine("\t- Pitch: " + fontInfo.Pitch);
                    Console.WriteLine("\t- Charset: " + fontInfo.Charset);
                    Console.WriteLine("\t- Panose:");
                    Console.WriteLine("\t\tFamily Kind: " + fontInfo.Panose[0]);
                    Console.WriteLine("\t\tSerif Style: " + fontInfo.Panose[1]);
                    Console.WriteLine("\t\tWeight: " + fontInfo.Panose[2]);
                    Console.WriteLine("\t\tProportion: " + fontInfo.Panose[3]);
                    Console.WriteLine("\t\tContrast: " + fontInfo.Panose[4]);
                    Console.WriteLine("\t\tStroke Variation: " + fontInfo.Panose[5]);
                    Console.WriteLine("\t\tArm Style: " + fontInfo.Panose[6]);
                    Console.WriteLine("\t\tLetterform: " + fontInfo.Panose[7]);
                    Console.WriteLine("\t\tMidline: " + fontInfo.Panose[8]);
                    Console.WriteLine("\t\tX-Height: " + fontInfo.Panose[9]);
                }
            }
            //ExEnd

            Assert.AreEqual(new int[] { 2, 15, 5, 2, 2, 2, 4, 3, 2, 4 }, doc.FontInfos["Calibri"].Panose);
            Assert.AreEqual(new int[] { 2, 2, 6, 3, 5, 4, 5, 2, 3, 4 }, doc.FontInfos["Times New Roman"].Panose);
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
            //ExSummary:Shows how to create a file font source.
            Document doc = new Document();

            // Create a font settings object for our document
            doc.FontSettings = new FontSettings();

            // Create a font source from a file in our system
            FileFontSource fileFontSource = new FileFontSource(MyDir + "Alte DIN 1451 Mittelschrift.ttf", 0);

            // Import the font source into our document
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
            //ExSummary:Shows how to create a folder font source.
            Document doc = new Document();

            // Create a font settings object for our document
            doc.FontSettings = new FontSettings();

            // Create a font source from a folder that contains font files
            FolderFontSource folderFontSource = new FolderFontSource(MyDir + "MyFonts", false, 1);

            // Add that source to our document
            doc.FontSettings.SetFontsSources(new FontSourceBase[] { folderFontSource });

            Assert.AreEqual(MyDir + "MyFonts", folderFontSource.FolderPath);
            Assert.AreEqual(false, folderFontSource.ScanSubfolders);
            Assert.AreEqual(FontSourceType.FontsFolder, folderFontSource.Type);
            Assert.AreEqual(1, folderFontSource.Priority);
            //ExEnd
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
            //ExSummary:Shows how to create a memory font source.
            Document doc = new Document();

            // Create a font settings object for our document
            doc.FontSettings = new FontSettings();

            // Import a font file, putting its contents into a byte array
            byte[] fontBytes = File.ReadAllBytes(MyDir + "Alte DIN 1451 Mittelschrift.ttf");

            // Create a memory font source from our array
            MemoryFontSource memoryFontSource = new MemoryFontSource(fontBytes, 0);

            // Add that font source to our document
            doc.FontSettings.SetFontsSources(new FontSourceBase[] { memoryFontSource });

            Assert.AreEqual(52208, memoryFontSource.FontData.Length);
            Assert.AreEqual(FontSourceType.MemoryFont, memoryFontSource.Type);
            Assert.AreEqual(0, memoryFontSource.Priority);
            //ExEnd
        }

        [Test]
        public void FontSourceSystem()
        {
            //ExStart
            //ExFor:Fonts.FontSettings.AddFontSubstitutes(String, String[])
            //ExFor:Fonts.FontSettings.EnableFontSubstitution
            //ExFor:Fonts.FontSettings.GetFontSubstitutes(String)
            //ExFor:Fonts.FontSettings.ResetFontSources
            //ExFor:Fonts.SystemFontSource
            //ExFor:Fonts.SystemFontSource.#ctor
            //ExFor:Fonts.SystemFontSource.#ctor(Int32)
            //ExFor:Fonts.SystemFontSource.GetSystemFontFolders
            //ExFor:Fonts.SystemFontSource.Type
            //ExSummary:Shows how to access a document's system font source and set font substitutes.
            Document doc = new Document();

            // Create a font settings object for our document
            doc.FontSettings = new FontSettings();

            // By default we always start with a system font source
            Assert.AreEqual(1, doc.FontSettings.GetFontsSources().Length);

            SystemFontSource systemFontSource = (SystemFontSource)doc.FontSettings.GetFontsSources()[0];
            Assert.AreEqual(FontSourceType.SystemFonts, systemFontSource.Type);
            Assert.AreEqual(0, systemFontSource.Priority);
            Assert.AreEqual(new[] { @"C:\WINDOWS\Fonts" }, SystemFontSource.GetSystemFontFolders());

            // Set a font that exists in the windows fonts directory as a substitute for one that doesn't
            doc.FontSettings.EnableFontSubstitution = true;
            doc.FontSettings.AddFontSubstitutes("Kreon-Regular", new string[] { "Calibri" });

            Assert.AreEqual(1, doc.FontSettings.GetFontSubstitutes("Kreon-Regular").Length);
            Assert.Contains("Calibri", doc.FontSettings.GetFontSubstitutes("Kreon-Regular"));

            // Alternatively, we could add a folder font source in which the corresponding folder contains the font
            FolderFontSource folderFontSource = new FolderFontSource(MyDir + "MyFonts", false);
            doc.FontSettings.SetFontsSources(new FontSourceBase[] { systemFontSource, folderFontSource });
            Assert.AreEqual(2, doc.FontSettings.GetFontsSources().Length);

            // Resetting the font sources still leaves us with the system font source as well as our substitutes
            doc.FontSettings.ResetFontSources();

            Assert.AreEqual(1, doc.FontSettings.GetFontsSources().Length);
            Assert.AreEqual(FontSourceType.SystemFonts, doc.FontSettings.GetFontsSources()[0].Type);
            Assert.AreEqual(1, doc.FontSettings.GetFontSubstitutes("Kreon-Regular").Length);
            //ExEnd
        }

        [Test]
        public void LoadFontFallbackSettingsFromFile()
        {
            //ExStart
            //ExFor:FontFallbackSettings.Load(String)
            //ExFor:FontFallbackSettings.Save(String)
            //ExSummary:Shows how to load and save font fallback settings from file.
            Document doc = new Document(MyDir + "Rendering.doc");
            
            // By default fallback settings are initialized with predefined settings which mimics the Microsoft Word fallback.
            FontSettings fontSettings = new FontSettings();
            fontSettings.FallbackSettings.Load(MyDir + "Fallback.xml");

            doc.FontSettings = fontSettings;
            doc.Save(MyDir + @"\Artifacts\LoadFontFallbackSettingsFromFile.pdf");

            // Saves font fallback setting by string
            doc.FontSettings.FallbackSettings.Save(MyDir + @"\Artifacts\FallbackSettings.xml");
            //ExEnd
        }

        [Test]
        public void LoadFontFallbackSettingsFromStream()
        {
            //ExStart
            //ExFor:FontFallbackSettings.Load(Stream)
            //ExFor:FontFallbackSettings.Save(Stream)
            //ExSummary:Shows how to load and save font fallback settings from stream.
            Document doc = new Document(MyDir + "Rendering.doc");

            // By default fallback settings are initialized with predefined settings which mimics the Microsoft Word fallback.
            using (FileStream fontFallbackStream = new FileStream(MyDir + "Fallback.xml", FileMode.Open))
            {
                FontSettings fontSettings = new FontSettings();
                fontSettings.FallbackSettings.Load(fontFallbackStream);

                doc.FontSettings = fontSettings;
            }

            doc.Save(MyDir + @"\Artifacts\LoadFontFallbackSettingsFromStream.pdf");

            // Saves font fallback setting by stream
            using (FileStream fontFallbackStream =
                new FileStream(MyDir + @"\Artifacts\FallbackSettings.xml", FileMode.Create))
            {
                doc.FontSettings.FallbackSettings.Save(fontFallbackStream);
            }
            //ExEnd
        }
    }
}