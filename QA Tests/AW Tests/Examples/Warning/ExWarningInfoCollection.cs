// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using NUnit.Framework;
using QA_Tests.Tests;

namespace QA_Tests.Examples.Warning
{
    [TestFixture]
    public class ExWarningInfoCollection : QaTestsBase
    {
        [Test]
        public void GetEnumeratorEx()
        {
            //ExStart
            //ExFor:GetEnumerator
            //ExId:GetEnumeratorEx
            //ExSummary:Shows how to use GetEnumerator.
            Aspose.Words.WarningInfoCollection wic = new Aspose.Words.WarningInfoCollection();
            var enumerator = wic.GetEnumerator();
            //ExEnd
        }
    }
}