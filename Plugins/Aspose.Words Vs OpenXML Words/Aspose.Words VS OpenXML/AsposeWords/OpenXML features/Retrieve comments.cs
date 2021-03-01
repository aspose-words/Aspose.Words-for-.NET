// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class RetrieveComments : TestUtil
    {
        [Test]
        public static void RetrieveCommentsFeature()
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(MyDir + "Comments.docx", false))
            {
                WordprocessingCommentsPart commentsPart = wordDoc.MainDocumentPart.WordprocessingCommentsPart;

                if (commentsPart?.Comments != null)
                    foreach (Comment comment in commentsPart.Comments.Elements<Comment>())
                        Console.WriteLine(comment.InnerText);
            }
        }
    }
}
