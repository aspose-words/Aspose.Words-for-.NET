// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System;
using System.Linq;
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
        public void InsertACommentFeature()
        {
            using (WordprocessingDocument document =
                WordprocessingDocument.Create(ArtifactsDir + "Insert a comment - OpenXML.docx",
                    WordprocessingDocumentType.Document))
            {
                // Locate the first paragraph in the document.
                Paragraph firstParagraph =
                    document.MainDocumentPart.Document.Descendants<Paragraph>().First();
                Comments comments;
                string id = "0";

                // Verify that the document contains a 
                // WordProcessingCommentsPart part; if not, add a new one.
                if (document.MainDocumentPart.GetPartsOfType<WordprocessingCommentsPart>().Any())
                {
                    comments =
                        document.MainDocumentPart.WordprocessingCommentsPart.Comments;
                    if (comments.HasChildren)
                        // Obtain an unused ID.
                        id = comments.Descendants<Comment>().Select(e => e.Id.Value).Max();
                }
                else
                {
                    // No "WordprocessingCommentsPart" part exists, so add one to the package.
                    WordprocessingCommentsPart commentPart =
                        document.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
                    commentPart.Comments = new Comments();
                    comments = commentPart.Comments;
                }

                // Compose a new Comment and add it to the Comments part.
                Paragraph p = new Paragraph(new Run(new Text("This is my comment.")));
                Comment cmt = new Comment
                {
                    Id = id,
                    Author = "author",
                    Initials = "initials",
                    Date = DateTime.Now
                };
                cmt.AppendChild(p);
                comments.AppendChild(cmt);
                comments.Save();

                // Specify the text range for the Comment. 
                // Insert the new CommentRangeStart before the first run of paragraph.
                firstParagraph.InsertBefore(new CommentRangeStart {Id = id}, firstParagraph.GetFirstChild<Run>());

                // Insert the new CommentRangeEnd after last run of paragraph.
                var cmtEnd = firstParagraph.InsertAfter(new CommentRangeEnd {Id = id},
                    firstParagraph.Elements<Run>().Last());

                // Compose a run with CommentReference and insert it.
                firstParagraph.InsertAfter(new Run(new CommentReference {Id = id}), cmtEnd);
            }
        }
    }
}
