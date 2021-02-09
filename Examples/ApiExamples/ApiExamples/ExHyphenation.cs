// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Aspose.Pdf.Text;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExHyphenation : ApiExampleBase
    {
        [Test]
        public void Dictionary()
        {
            //ExStart
            //ExFor:Hyphenation.IsDictionaryRegistered(String)
            //ExFor:Hyphenation.RegisterDictionary(String, String)
            //ExFor:Hyphenation.UnregisterDictionary(String)
            //ExSummary:Shows how to register a hyphenation dictionary.
            // A hyphenation dictionary contains a list of strings that define hyphenation rules for the dictionary's language.
            // When a document contains lines of text in which a word could be split up and continued on the next line,
            // hyphenation will look through the dictionary's list of strings for that word's substrings.
            // If the dictionary contains a substring, then hyphenation will split the word across two lines
            // by the substring and add a hyphen to the first half.
            // Register a dictionary file from the local file system to the "de-CH" locale.
            Hyphenation.RegisterDictionary("de-CH", MyDir + "hyph_de_CH.dic");

            Assert.True(Hyphenation.IsDictionaryRegistered("de-CH"));
            
            // Open a document containing text with a locale matching that of our dictionary,
            // and save it to a fixed-page save format. The text in that document will be hyphenated.
            Document doc = new Document(MyDir + "German text.docx");

            Assert.True(doc.FirstSection.Body.FirstParagraph.Runs.OfType<Run>().All(
                r => r.Font.LocaleId == new CultureInfo("de-CH").LCID));

            doc.Save(ArtifactsDir + "Hyphenation.Dictionary.Registered.pdf");

            // Re-load the document after un-registering the dictionary,
            // and save it to another PDF, which will not have hyphenated text.
            Hyphenation.UnregisterDictionary("de-CH");

            Assert.False(Hyphenation.IsDictionaryRegistered("de-CH"));

            doc = new Document(MyDir + "German text.docx");
            doc.Save(ArtifactsDir + "Hyphenation.Dictionary.Unregistered.pdf");
            //ExEnd

#if NET462 || NETCOREAPP2_1 || JAVA
            Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "Hyphenation.Dictionary.Registered.pdf");
            TextAbsorber textAbsorber = new TextAbsorber();
            textAbsorber.Visit(pdfDoc);

            Assert.True(textAbsorber.Text.Contains("La ob storen an deinen am sachen. Dop-\r\n" +
                                                   "pelte  um  da  am  spateren  verlogen  ge-\r\n" +
                                                   "kommen  achtzehn  blaulich."));

            pdfDoc = new Aspose.Pdf.Document(ArtifactsDir + "Hyphenation.Dictionary.Unregistered.pdf");
            textAbsorber = new TextAbsorber();
            textAbsorber.Visit(pdfDoc);

            Assert.True(textAbsorber.Text.Contains("La  ob  storen  an  deinen  am  sachen. \r\n" +
                                                   "Doppelte  um  da  am  spateren  verlogen \r\n" +
                                                   "gekommen  achtzehn  blaulich."));
#endif
        }

        //ExStart
        //ExFor:Hyphenation
        //ExFor:Hyphenation.Callback
        //ExFor:Hyphenation.RegisterDictionary(String, Stream)
        //ExFor:Hyphenation.RegisterDictionary(String, String)
        //ExFor:Hyphenation.WarningCallback
        //ExFor:IHyphenationCallback
        //ExFor:IHyphenationCallback.RequestDictionary(System.String)
        //ExSummary:Shows how to open and register a dictionary from a file.
        [Test] //ExSkip
        public void RegisterDictionary()
        {
            // Set up a callback that tracks warnings that occur during hyphenation dictionary registration.
            WarningInfoCollection warningInfoCollection = new WarningInfoCollection();
            Hyphenation.WarningCallback = warningInfoCollection;

            // Register an English (US) hyphenation dictionary by stream.
            Stream dictionaryStream = new FileStream(MyDir + "hyph_en_US.dic", FileMode.Open);
            Hyphenation.RegisterDictionary("en-US", dictionaryStream);

            Assert.AreEqual(0, warningInfoCollection.Count);

            // Open a document with a locale that Microsoft Word may not hyphenate on an English machine, such as German.
            Document doc = new Document(MyDir + "German text.docx");

            // To hyphenate that document upon saving, we need a hyphenation dictionary for the "de-CH" language code.
            // This callback will handle the automatic request for that dictionary.
            Hyphenation.Callback = new CustomHyphenationDictionaryRegister();

            // When we save the document, German hyphenation will take effect.
            doc.Save(ArtifactsDir + "Hyphenation.RegisterDictionary.pdf");

            // This dictionary contains two identical patterns, which will trigger a warning.
            Assert.AreEqual(1, warningInfoCollection.Count);
            Assert.AreEqual(WarningType.MinorFormattingLoss, warningInfoCollection[0].WarningType);
            Assert.AreEqual(WarningSource.Layout, warningInfoCollection[0].Source);
            Assert.AreEqual("Hyphenation dictionary contains duplicate patterns. The only first found pattern will be used. " +
                            "Content can be wrapped differently.", warningInfoCollection[0].Description);
        }

        /// <summary>
        /// Associates ISO language codes with local system filenames for hyphenation dictionary files.
        /// </summary>
        private class CustomHyphenationDictionaryRegister : IHyphenationCallback
        {
            public CustomHyphenationDictionaryRegister()
            {
                mHyphenationDictionaryFiles = new Dictionary<string, string>
                {
                    { "en-US", MyDir + "hyph_en_US.dic" },
                    { "de-CH", MyDir + "hyph_de_CH.dic" }
                };
            }

            public void RequestDictionary(string language)
            {
                Console.Write("Hyphenation dictionary requested: " + language);

                if (Hyphenation.IsDictionaryRegistered(language))
                {
                    Console.WriteLine(", is already registered.");
                    return;
                }

                if (mHyphenationDictionaryFiles.ContainsKey(language))
                {
                    Hyphenation.RegisterDictionary(language, mHyphenationDictionaryFiles[language]);
                    Console.WriteLine(", successfully registered.");
                    return;
                }

                Console.WriteLine(", no respective dictionary file known by this Callback.");
            }

            private readonly Dictionary<string, string> mHyphenationDictionaryFiles;
        }
        //ExEnd
    }
}