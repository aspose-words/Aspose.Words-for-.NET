// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using NUnit.Framework;
using System;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class InsertAComment : TestUtil
    {
        [Test]
        public void InsertCommentAsposeWords()
        {
            //ExStart:InsertCommentAsposeWords
            //GistId:2ee1b3b132932dacb41cdc51fdc03888
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Comment newComment = new Comment(doc)
            {
                Author = "Aspose.Words",
                Initial = "AW",
                DateTime = DateTime.Now
            };
            newComment.SetText("Comment regarding text.");

            // Add text to the document, warp it in a comment range, and then add your comment.
            Paragraph para = doc.FirstSection.Body.FirstParagraph;
            para.AppendChild(new CommentRangeStart(doc, newComment.Id));
            para.AppendChild(new Run(doc, "Commented text."));
            para.AppendChild(new CommentRangeEnd(doc, newComment.Id));
            para.AppendChild(newComment);

            doc.Save(ArtifactsDir + "Insert comment - Aspose.Words.docx");
            //ExEnd:InsertCommentAsposeWords
        }
    }
}
