// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.Border
{
    [TestFixture]
    public class ExBorderCollection : QaTestsBase
    {
        [Test]
        public void GetEnumeratorEx()
        {
            //ExStart
            //ExFor:GetEnumerator
            //ExId:GetEnumeratorEx
            //ExSummary:Shows how to use GetEnumerator.
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder();
            Aspose.Words.BorderCollection borders = builder.ParagraphFormat.Borders;

            var enumerator = borders.GetEnumerator();
            //ExEnd
        }

        [Test]
        public void ClearFormattingEx()
        {
            //ExStart
            //ExFor:ClearFormatting
            //ExId:ClearFormattingEx
            //ExSummary:Shows how to use ClearFormatting.
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder();
            Aspose.Words.BorderCollection borders = builder.ParagraphFormat.Borders;

            borders.ClearFormatting();
            //ExEnd
        }
    }
}
