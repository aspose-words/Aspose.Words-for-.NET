// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
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

            // Register by string
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
            //ExFor:Hyphenation.IsDictionaryRegistered(string)
            //ExSummary:Shows how to open check if some dictionary is registered.
            Document doc = new Document(MyDir + "Document.doc");
            Hyphenation.RegisterDictionary("en-US", MyDir + "hyph_en_US.dic");

            Console.WriteLine(Hyphenation.IsDictionaryRegistered("en-US")); // True
            //ExEnd
        }

        [Test]
        public void UnregisterDictionaryEx()
        {
            //ExStart
            //ExFor:Hyphenation.UnregisterDictionary(string)
            //ExSummary:Shows how to un-register a dictionary
            Document doc = new Document(MyDir + "Document.doc");
            Hyphenation.RegisterDictionary("en-US", MyDir + "hyph_en_US.dic");

            Hyphenation.UnregisterDictionary("en-US");

            Console.WriteLine(Hyphenation.IsDictionaryRegistered("en-US")); // False
            //ExEnd
        }
    }
}
