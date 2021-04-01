// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class DeleteCommentsByAllOrASpecificAuthor : TestUtil
    {
        [Test]
        public void DeleteCommentsByAllOrASpecificAuthorFeature()
        {
            RemoveComments("");
        }

        private void RemoveComments(string author)
        {
            // Get an existing Wordprocessing document.
            using (WordprocessingDocument document =
                WordprocessingDocument.Open(MyDir + "Comments.docx", true))
            {
                WordprocessingCommentsPart commentPart =
                    document.MainDocumentPart.WordprocessingCommentsPart;

                // If no "WordprocessingCommentsPart" exists, there can be no comments.
                // Stop execution and return from the method.
                if (commentPart == null)
                    return;

                // Create a list of comments by the specified author.
                // If the author name is empty, then list all authors.
                List<Comment> commentsToDelete =
                    commentPart.Comments.Elements<Comment>().ToList();

                if (!String.IsNullOrEmpty(author))
                {
                    commentsToDelete = commentsToDelete.Where(c => c.Author == author).ToList();
                }

                IEnumerable<string> commentIds =
                    commentsToDelete.Select(r => r.Id.Value);

                foreach (Comment c in commentsToDelete)
                    c.Remove();

                // Save changes to the comments part.
                commentPart.Comments.Save();

                Document doc = document.MainDocumentPart.Document;

                // Delete the "CommentRangeStart" for each deleted comment in the main document.
                List<CommentRangeStart> commentRangeStartToDelete =
                    doc.Descendants<CommentRangeStart>().Where(c => commentIds.Contains(c.Id.Value)).ToList();

                foreach (CommentRangeStart c in commentRangeStartToDelete)
                    c.Remove();

                // Delete the "CommentRangeEnd" for each deleted comment in the main document.
                List<CommentRangeEnd> commentRangeEndToDelete =
                    doc.Descendants<CommentRangeEnd>().Where(c => commentIds.Contains(c.Id.Value)).ToList();

                foreach (CommentRangeEnd c in commentRangeEndToDelete)
                    c.Remove();

                // Delete the "CommentReference" for each deleted comment in the main document.
                List<CommentReference> commentRangeReferenceToDelete =
                    doc.Descendants<CommentReference>().Where(c => commentIds.Contains(c.Id.Value)).ToList();

                foreach (CommentReference c in commentRangeReferenceToDelete)
                    c.Remove();

                using (Stream stream = File.Open(ArtifactsDir + "Remove comments - OpenXML.docx", FileMode.CreateNew))
                {
                    doc.Save(stream);
                }

            }
        }
    }
}
