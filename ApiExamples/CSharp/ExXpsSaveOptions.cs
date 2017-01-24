// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Saving;

using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExXpsSaveOptions : ApiExampleBase
    {
        [Test]
        public void OptimizeOutput()
        {
            Document doc = new Document(MyDir + "XPSOutputOptimize.docx");

            XpsSaveOptions saveOptions = new XpsSaveOptions();
            saveOptions.OptimizeOutput = true;

            doc.Save(MyDir + @"\Artifacts\XPSOutputOptimize.docx");
        }
    }
}
