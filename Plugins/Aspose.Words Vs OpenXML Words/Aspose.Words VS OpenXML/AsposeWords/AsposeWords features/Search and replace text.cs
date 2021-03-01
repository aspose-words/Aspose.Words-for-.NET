// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System.Text.RegularExpressions;
using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class SearchAndReplaceText : TestUtil
    {
        [Test]
        public static void SearchAndReplaceTextFeature()
        {
            Document doc = new Document(MyDir + "Search and replace text.docx");

            Regex regex = new Regex("Hello World!", RegexOptions.IgnoreCase);
            doc.Range.Replace(regex, "Hi Everyone!");

            doc.Save(ArtifactsDir + "Search and replace text - Aspose.Words.docx");
        }
    }
}
