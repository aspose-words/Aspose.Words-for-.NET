// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
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
        //ExFor:IHyphenationCallback
        //ExFor:IHyphenationCallback.RequestDictionary(System.String)
        //ExSummary:Shows how to open and register a dictionary from a file.
        [Test] //ExSKip
        public void RegisterDictionaryEx()
        {
            Document doc = new Document(MyDir + "Document.doc");

            // Register by stream
            Stream dictionaryStream = new FileStream(MyDir + "hyph_de_CH.dic", FileMode.Open);
            Hyphenation.RegisterDictionary("de-CH", dictionaryStream);

            // Register by string via callback
            Hyphenation.WarningCallback = new HyphenationWarnings();

            Hyphenation.Callback = new HyphenationPrinter();
            Hyphenation.Callback.RequestDictionary("en-US");
        }

        /// <summary>
        /// Associates ISO language codes with dictionary files for their respective languages
        /// </summary>
        private class HyphenationPrinter : IHyphenationCallback
        {
            public HyphenationPrinter()
            {
                mDictionaryFilenames = new Dictionary<string, string>
                {
                    { "en-US", MyDir + "hyph_en_US.dic" },
                    { "de-CH", MyDir + "hyph_de_CH.dic" }
                };
            }

            public void RequestDictionary(string language)
            {
                if (mDictionaryFilenames.ContainsKey(language) && !Hyphenation.IsDictionaryRegistered(language))
                {
                    Hyphenation.RegisterDictionary(language, mDictionaryFilenames[language]);
                }
            }

            private readonly Dictionary<string, string> mDictionaryFilenames;
        }

        /// <summary>
        /// Prints hyphenation warnings
        /// </summary>
        private class HyphenationWarnings : IWarningCallback
        {
            void IWarningCallback.Warning(WarningInfo info)
            {
                Console.WriteLine(info.Description);
            }
        }
        //ExEnd

        [Test]
        public void IsDictionaryRegisteredEx()
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
        public void UnregisteredDictionaryEx()
        {
            //ExStart
            //ExFor:Hyphenation.UnregisterDictionary(String)
            //ExSummary:Shows how to un-register a dictionary
            Document doc = new Document(MyDir + "Document.doc");
            Hyphenation.RegisterDictionary("en-US", MyDir + "hyph_en_US.dic");

            Hyphenation.UnregisterDictionary("en-US");

            Assert.AreEqual(false, Hyphenation.IsDictionaryRegistered("en-US"));
            //ExEnd
        }
    }
}