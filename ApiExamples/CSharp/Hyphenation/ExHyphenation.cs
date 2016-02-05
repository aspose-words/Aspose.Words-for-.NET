// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;


namespace ApiExamples.Hyphenation
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
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");

            // Register by string
            Aspose.Words.Hyphenation.RegisterDictionary("en-US", MyDir + @"hyph_en_US.dic");

            // Register by stream
            Stream dictionaryStream = new FileStream(MyDir + @"hyph_de_CH.dic", FileMode.Open);
            Aspose.Words.Hyphenation.RegisterDictionary("de-CH", dictionaryStream);
            //ExEnd
        }

        [Test]
        public void IsDictionaryRegisteredEx()
        {
            //ExStart
            //ExFor:Hyphenation.IsDictionaryRegistered(string)
            //ExSummary:Shows how to open check if some dictionary is registered.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");
            Aspose.Words.Hyphenation.RegisterDictionary("en-US", MyDir + @"hyph_en_US.dic");

            Console.WriteLine(Aspose.Words.Hyphenation.IsDictionaryRegistered("en-US")); // True
            //ExEnd
        }

        [Test]
        public void UnregisterDictionaryEx()
        {
            //ExStart
            //ExFor:Hyphenation.UnregisterDictionary(string)
            //ExSummary:Shows how to un-register a dictionary
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.doc");
            Aspose.Words.Hyphenation.RegisterDictionary("en-US", MyDir + @"hyph_en_US.dic");

            Aspose.Words.Hyphenation.UnregisterDictionary("en-US");

            Console.WriteLine(Aspose.Words.Hyphenation.IsDictionaryRegistered("en-US")); // False
            //ExEnd
        }
    }
}
