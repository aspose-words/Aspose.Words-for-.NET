//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using Aspose.Words;
using NUnit.Framework;

namespace Examples
{
    /// <summary>
    /// Examples for the .NET vs Java Differences in Aspose.Words in the Programmers Guide.
    /// </summary>
    [TestFixture]
    public class ExDotNetVsJava
    {
        //ExStart
        //ExId:SaveSignature
        //ExSummary:Shows difference in .NET and Java in signatures of a method with an enum parameter.
        // The saveFormat parameter is a SaveFormat enum value.
        void Save(string fileName, SaveFormat saveFormat)
        //ExEnd
        {
            // Do nothing.
        }

        //ExStart
        //ExId:CollectionItemSignature
        //ExSummary:Shows difference in signatures of collection indexers in .NET vs Java.
        public class HeaderFooterCollection
        {
            // Get by index is an indexer.
            public HeaderFooter this[int index]
            {                       //ExSkip
                get                 //ExSkip
                {                   //ExSkip
                    return null;    //ExSkip
                }                   //ExSkip
            }                       //ExSkip

            // Get by header footer type is an overloaded indexer.
            public HeaderFooter this[HeaderFooterType headerFooterType]
            {                       //ExSkip
                get                 //ExSkip
                {                   //ExSkip
                    return null;    //ExSkip
                }                   //ExSkip
            }                       //ExSkip
        }
        //ExEnd
    }
}
