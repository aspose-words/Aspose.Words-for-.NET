// Copyright (c) 2001-2019 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExHyphenation : ApiExampleBase
    {
        [Test]
        public void RegisterDictionaryEx()
        {
            //ExStart
            //ExFor:Hyphenation.RegisterDictionary(String, Stream)
            //ExFor:Hyphenation.RegisterDictionary(String, String)
            //ExSummary:Shows how to open and register a dictionary from a file.
            Document doc = new Document(MyDir + "Document.doc");

            // Register by String
            Hyphenation.RegisterDictionary("en-US", MyDir + "hyph_en_US.dic");

            // Register by stream
            Stream dictionaryStream = new FileStream(MyDir + "hyph_de_CH.dic", FileMode.Open);
            Hyphenation.RegisterDictionary("de-CH", dictionaryStream);
            //ExEnd
        }

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