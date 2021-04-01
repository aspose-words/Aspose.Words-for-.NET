// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class RemovePageBreaks : TestUtil
    {
        [Test]
        public void RemovePageBreaksFeature()
        {
            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(MyDir + "Remove page breaks.docx", true))
            {
                MainDocumentPart mainPart = myDoc.MainDocumentPart;
                List<Break> breaks = mainPart.Document.Descendants<Break>().ToList();

                foreach (Break b in breaks)
                    b.Remove();

                using (Stream stream = File.Create(ArtifactsDir + "Remove page breaks - OpenXML.docx"))
                {
                    mainPart.Document.Save(stream);
                }
            }
        }
    }
}
