// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.IO;
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
            //ExSummary:Shows how to perform and verify hyphenation dictionary registration.
            // Register a dictionary file from the local file system to the "de-CH" locale
            Hyphenation.RegisterDictionary("de-CH", MyDir + "hyph_de_CH.dic");

            // This method can be used to verify that a language has a matching registered hyphenation dictionary
            Assert.True(Hyphenation.IsDictionaryRegistered("de-CH"));

            // The dictionary file contains a long list of words in a specified language, and in this case it is German
            // These words define a set of rules for hyphenating text (splitting words across lines)
            // If we open a document with text of a language matching that of a registered dictionary,
            // that dictionary's hyphenation rules will be applied and visible upon saving
            Document doc = new Document(MyDir + "German text.docx");
            doc.Save(ArtifactsDir + "Hyphenation.Dictionary.Registered.pdf");

            // We can also un-register a dictionary to disable these effects on any documents opened after the operation
            Hyphenation.UnregisterDictionary("de-CH");

            Assert.False(Hyphenation.IsDictionaryRegistered("de-CH"));

            doc = new Document(MyDir + "German text.docx");
            doc.Save(ArtifactsDir + "Hyphenation.Dictionary.Unregistered.pdf");
            //ExEnd
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
            // Set up a callback that tracks warnings that occur during hyphenation dictionary registration
            WarningInfoCollection warningInfoCollection = new WarningInfoCollection();
            Hyphenation.WarningCallback = warningInfoCollection;

            // Register an English (US) hyphenation dictionary by stream
            Stream dictionaryStream = new FileStream(MyDir + "hyph_en_US.dic", FileMode.Open, FileAccess.Read);
            Hyphenation.RegisterDictionary("en-US", dictionaryStream);

            // No warnings detected
            Assert.AreEqual(0, warningInfoCollection.Count);

            // Open a document with a German locale that might not get automatically hyphenated by Microsoft Word an english machine
            Document doc = new Document(MyDir + "German text.docx");

            // To hyphenate that document upon saving, we need a hyphenation dictionary for the "de-CH" language code
            // This callback will handle the automatic request for that dictionary 
            Hyphenation.Callback = new CustomHyphenationDictionaryRegister();

            // When we save the document, it will be hyphenated according to rules defined by the dictionary known by our callback
            doc.Save(ArtifactsDir + "Hyphenation.RegisterDictionary.pdf");

            // This dictionary contains two identical patterns, which will trigger a warning
            Assert.AreEqual(1, warningInfoCollection.Count);
            Assert.AreEqual(WarningType.MinorFormattingLoss, warningInfoCollection[0].WarningType);
            Assert.AreEqual(WarningSource.Layout, warningInfoCollection[0].Source);
            Assert.AreEqual("Hyphenation dictionary contains duplicate patterns. The only first found pattern will be used. " +
                            "Content can be wrapped differently.", warningInfoCollection[0].Description);
        }

        /// <summary>
        /// Associates ISO language codes with custom local system dictionary files for their respective languages
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