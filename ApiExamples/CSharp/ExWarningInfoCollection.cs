// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;

using Aspose.Words;

using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExWarningInfoCollection : ApiExampleBase
    {
        [Test]
        public void GetEnumeratorEx()
        {
            //ExStart
            //ExFor:WarningInfoCollection.GetEnumerator
            //ExFor:WarningInfoCollection.Clear
            //ExSummary:Shows how to read and clear a collection of warnings.
            WarningInfoCollection wic = new WarningInfoCollection();

            var enumerator = wic.GetEnumerator();
            while (enumerator.MoveNext())
            {
                WarningInfo wi = (WarningInfo)enumerator.Current;
                Console.WriteLine(wi.Description);
            }

            wic.Clear();
            //ExEnd
        }
    }
}