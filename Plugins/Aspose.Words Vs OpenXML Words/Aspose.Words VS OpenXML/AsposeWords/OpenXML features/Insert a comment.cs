// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class InsertAComment : TestUtil
    {
        [Test]
        public void InsertComment()
        {
            using WordprocessingDocument doc = WordprocessingDocument.Create(ArtifactsDir + "Insert a comment - OpenXML.docx", 
                WordprocessingDocumentType.Document);

            // Add the main document part.
            MainDocumentPart mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = new Body();

            // Add a paragraph with some text.
            Paragraph paragraph = new Paragraph();
            Run run = new Run();
            run.AppendChild(new Text("Commented text."));
            paragraph.AppendChild(run);
            body.AppendChild(paragraph);

            // Add a comments part to the document.
            WordprocessingCommentsPart commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>();
            commentsPart.Comments = new Comments();

            // Create a comment.
            Comment comment = new Comment()
            {
                Id = "1",
                Author = "Aspose.Words",
                Date = DateTime.Now
            };

            // Add text to the comment.
            Paragraph commentParagraph = new Paragraph();
            Run commentRun = new Run();
            commentRun.AppendChild(new Text("Comment regarding text."));
            commentParagraph.AppendChild(commentRun);
            comment.AppendChild(commentParagraph);

            // Add the comment to the comments part.
            commentsPart.Comments.AppendChild(comment);
            commentsPart.Comments.Save();

            // Add a reference to the comment in the document.
            run.AppendChild(new CommentRangeStart { Id = "1" });
            run.AppendChild(new CommentRangeEnd { Id = "1" });
            run.AppendChild(new CommentReference { Id = "1" });

            // Add the body to the document.
            mainPart.Document.AppendChild(body);

            // Save the document.
            mainPart.Document.Save();
        }
    }
}
