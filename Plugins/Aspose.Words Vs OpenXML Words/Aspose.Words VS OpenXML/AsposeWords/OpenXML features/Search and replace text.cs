// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System.IO;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class SearchAndReplaceText : TestUtil
    {
        [Test]
        public static void SearchAndReplaceTextFeature()
        {
            using (WordprocessingDocument wordDoc =
                WordprocessingDocument.Open(MyDir + "Search and replace text.docx", true))
            {
                string docText;

                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                Regex regexText = new Regex("Hello world!");
                docText = regexText.Replace(docText, "Hi Everyone!");

                using (StreamWriter sw =
                    new StreamWriter(File.Create(ArtifactsDir + "Search and replace text - OpenXML.docx")))
                {
                    sw.Write(docText);
                }
            }
        }
    }
}
