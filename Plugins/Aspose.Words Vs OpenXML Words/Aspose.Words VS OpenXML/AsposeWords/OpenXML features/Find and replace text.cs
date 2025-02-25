// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class FindAndReplaceText : TestUtil
    {
        [Test]
        public static void FindAndReplaceTextFeature()
        {
            File.Copy(MyDir + "Find and replace text.docx", ArtifactsDir + "Find and replace text - OpenXML.docx", true);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(ArtifactsDir + "Find and replace text - OpenXML.docx", true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(doc.MainDocumentPart.GetStream()))
                    docText = sr.ReadToEnd();

                Regex regexText = new Regex("Hello World!");
                docText = regexText.Replace(docText, "Hi Everyone!");

                using (StreamWriter sw = new StreamWriter(doc.MainDocumentPart.GetStream(FileMode.Create)))
                    sw.Write(docText);
            }
        }
    }
}
