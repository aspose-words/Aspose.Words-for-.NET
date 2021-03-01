// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features.Features_missing_in_OpenXML
{
    [TestFixture]
    public class FindAndReplace : TestUtil
    {
        [Test]
        public static void FindAndReplaceFeature()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("This one is sad.");
            builder.Writeln("That one is mad.");

            FindReplaceOptions options = new FindReplaceOptions();
            options.MatchCase = false;
            options.FindWholeWordsOnly = true;
            
            ReplaceOneWordWithAnother(doc);
            doc.Save(ArtifactsDir + "Find and replace.Replace one word - Aspose.Words.docx");

            ReplaceTwoSimilarWords(doc);
            doc.Save(ArtifactsDir + "Find and replace.Replace two words - Aspose.Words.docx");
        }

        static void ReplaceOneWordWithAnother(Document doc)
        {
            doc.Range.Replace("sad", "bad");
        }

        static void ReplaceTwoSimilarWords(Document doc)
        {
            doc.Range.Replace(new Regex("[s|m]ad"), "bad");
        }
    }
}
