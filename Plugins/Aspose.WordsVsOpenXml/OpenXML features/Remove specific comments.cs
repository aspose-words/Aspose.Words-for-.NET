// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

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
    public class RemoveSpecificComments : TestUtil
    {
        [Test]
        public void RemoveCommentsOpenXml()
        {
            //ExStart:RemoveCommentsOpenXml
            //GistId:787486ce8310219ee50379944022f5db
            string authorName = string.Empty;
            File.Copy(MyDir + "Comments.docx", ArtifactsDir + "Remove comments - OpenXML.docx", true);

            using WordprocessingDocument doc = WordprocessingDocument.Open(ArtifactsDir + "Remove comments - OpenXML.docx", true);

            // Add a main document part to the new document.
            MainDocumentPart mainPart = doc.MainDocumentPart;
            // Get the comments part
            WordprocessingCommentsPart commentsPart = mainPart.WordprocessingCommentsPart;
            if (commentsPart != null)
            {
                // Get the comments
                Comments comments = commentsPart.Comments;
                // Create a list to hold comments to remove
                List<Comment> commentsToRemove = new();

                foreach (var comment in comments.Elements<Comment>())
                    if (string.IsNullOrEmpty(authorName) || comment.Author == authorName)
                        commentsToRemove.Add(comment);

                IEnumerable<string> commentIds =
                commentsToRemove.Select(r => r.Id.Value);

                // Remove the comments
                foreach (Comment comment in commentsToRemove)
                    comment.Remove();

                // Save changes to the comments part
                commentsPart.Comments.Save();

                // Delete the "CommentRangeStart" for each deleted comment in the main document.
                List<CommentRangeStart> commentRangeStartToDelete =
                    mainPart.Document.Descendants<CommentRangeStart>().Where(c => commentIds.Contains(c.Id.Value)).ToList();

                foreach (CommentRangeStart rangeStart in commentRangeStartToDelete)
                    rangeStart.Remove();

                // Delete the "CommentRangeEnd" for each deleted comment in the main document.
                List<CommentRangeEnd> commentRangeEndToDelete =
                    mainPart.Document.Descendants<CommentRangeEnd>().Where(c => commentIds.Contains(c.Id.Value)).ToList();

                foreach (CommentRangeEnd rangeEnd in commentRangeEndToDelete)
                    rangeEnd.Remove();

                // Delete the "CommentReference" for each deleted comment in the main document.
                List<CommentReference> commentRangeReferenceToDelete =
                    mainPart.Document.Descendants<CommentReference>().Where(c => commentIds.Contains(c.Id.Value)).ToList();

                foreach (CommentReference reference in commentRangeReferenceToDelete)
                    reference.Remove();
                //ExEnd:RemoveCommentsOpenXml
            }
        }
    }
}
