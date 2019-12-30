// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
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

            using (IEnumerator<WarningInfo> enumerator = wic.GetEnumerator())
            {
                while (enumerator.MoveNext())
                {
                    WarningInfo wi = enumerator.Current;
                    if (wi != null) Console.WriteLine(wi.Description);
                }

                wic.Clear();
            }
            //ExEnd
        }
    }
}