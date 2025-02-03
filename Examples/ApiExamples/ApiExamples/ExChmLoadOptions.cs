// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExChmLoadOptions : ApiExampleBase
    {
        [Test] // ToDo: Need to add tests.
        public void OriginalFileName()
        {
            //ExStart
            //ExFor:ChmLoadOptions
            //ExFor:ChmLoadOptions.#ctor
            //ExFor:ChmLoadOptions.OriginalFileName
            //ExSummary:Shows how to resolve URLs like "ms-its:myfile.chm::/index.htm".
            // Our document contains URLs like "ms-its:amhelp.chm::....htm", but it has a different name,
            // so file links don't work after saving it to HTML.
            // We need to define the original filename in 'ChmLoadOptions' to avoid this behavior.
            ChmLoadOptions loadOptions = new ChmLoadOptions { OriginalFileName = "amhelp.chm" };

            Document doc = new Document(new MemoryStream(File.ReadAllBytes(MyDir + "Document with ms-its links.chm")),
                loadOptions);

            doc.Save(ArtifactsDir + "ExChmLoadOptions.OriginalFileName.html");
            //ExEnd
        }
    }
}
