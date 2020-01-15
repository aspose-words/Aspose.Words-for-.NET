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
            Stream dictionaryStream = new FileStream(MyDir + "hyph_en_US.dic", FileMode.Open);
            Hyphenation.RegisterDictionary("en-US", dictionaryStream);

            // No warnings detected
            Assert.AreEqual(0, warningInfoCollection.Count);

            // Open a document with a German locale that might not get automatically hyphenated by Microsoft Word an english machine
            Document doc = new Document(MyDir + "RandomGermanWords.doc");

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

        [Test]
        public void IsDictionaryRegistered()
        {
            //ExStart
            //ExFor:Hyphenation.IsDictionaryRegistered(String)
            //ExSummary:Shows how to open check if some dictionary is registered.
            Document doc = new Document(MyDir + "Document.doc");
            Hyphenation.RegisterDictionary("en-US", MyDir + "hyph_en_US.dic");

            Assert.AreEqual(true, Hyphenation.IsDictionaryRegistered("en-US"));
            //ExEnd
        }

        [Test]
        public void UnregisteredDictionary()
        {
            //ExStart
            //ExFor:Hyphenation.UnregisterDictionary(String)
            //ExSummary:Shows how to un-register a dictionary.
            Document doc = new Document(MyDir + "Document.doc");
            Hyphenation.RegisterDictionary("en-US", MyDir + "hyph_en_US.dic");

            Hyphenation.UnregisterDictionary("en-US");

            Assert.AreEqual(false, Hyphenation.IsDictionaryRegistered("en-US"));
            //ExEnd
        }
    }
}