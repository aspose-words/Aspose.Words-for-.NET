// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

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
            using WordprocessingDocument doc = WordprocessingDocument.Open(MyDir + "Comments.docx", false);

            WordprocessingCommentsPart commentsPart = doc.MainDocumentPart.WordprocessingCommentsPart;

            if (commentsPart?.Comments != null)
                foreach (Comment comment in commentsPart.Comments.Elements<Comment>())
                    Console.WriteLine(comment.InnerText);
        }
    }
}
