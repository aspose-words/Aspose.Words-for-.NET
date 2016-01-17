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
using QA_Tests.Tests;

namespace QA_Tests.Examples.Hyphenation
{
    [TestFixture]
    public class ExHyphenation : QaTestsBase
    {
        [Test]
        public void AllHyphenationEx()
        {
            //ExStart
            //ExFor:Hyphenation.RegisterDictionary(string, stream)
            //ExFor:Hyphenation.RegisterDictionary(string, string)
            //ExFor:Hyphenation.IsDictionaryRegistered(string)
            //ExFor:Hyphenation.UnregisterDictionary(string)
            //ExSummary:Shows how to open and register a dictionary from a file.
            Aspose.Words.Document doc = new Aspose.Words.Document(ExDir + "Document.doc");

            // Register by string
            Aspose.Words.Hyphenation.RegisterDictionary("en-US", ExDir + @"hyph_en_US.dic");

            // Confirm that en-US is registered
            Console.WriteLine(Aspose.Words.Hyphenation.IsDictionaryRegistered("en-US"));

            // Register by stream
            Stream dictionaryStream = new FileStream(ExDir + @"hyph_de_CH.dic", FileMode.Open);
            Aspose.Words.Hyphenation.RegisterDictionary("de-CH", dictionaryStream);

            // Unregister
            Aspose.Words.Hyphenation.UnregisterDictionary("en-US");

            // Confirm that en-US is no longer registered
            Console.WriteLine(Aspose.Words.Hyphenation.IsDictionaryRegistered("en-US"));
            //ExEnd
        }
    }
}
