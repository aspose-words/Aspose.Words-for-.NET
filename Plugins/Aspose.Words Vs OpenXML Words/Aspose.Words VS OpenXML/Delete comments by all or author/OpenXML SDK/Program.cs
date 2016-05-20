// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string File = FilePath + "Delete comments - OpenXML.docx";

            DeleteComments(File, "");
        }
        // Delete comments by a specific author. Pass an empty string for the 
        // author to delete all comments, by all authors.
        public static void DeleteComments(string fileName,
            string author = "")
        {
            // Get an existing Wordprocessing document.
            using (WordprocessingDocument document =
                WordprocessingDocument.Open(fileName, true))
            {
                // Set commentPart to the document WordprocessingCommentsPart, 
                // if it exists.
                WordprocessingCommentsPart commentPart =
                    document.MainDocumentPart.WordprocessingCommentsPart;

                // If no WordprocessingCommentsPart exists, there can be no 
                // comments. Stop execution and return from the method.
                if (commentPart == null)
                {
                    return;
                }

                // Create a list of comments by the specified author, or
                // if the author name is empty, all authors.
                List<Comment> commentsToDelete =
                    commentPart.Comments.Elements<Comment>().ToList();
                if (!String.IsNullOrEmpty(author))
                {
                    commentsToDelete = commentsToDelete.
                    Where(c => c.Author == author).ToList();
                }
                IEnumerable<string> commentIds =
                    commentsToDelete.Select(r => r.Id.Value);

                // Delete each comment in commentToDelete from the 
                // Comments collection.
                foreach (Comment c in commentsToDelete)
                {
                    c.Remove();
                }

                // Save the comment part change.
                commentPart.Comments.Save();

                Document doc = document.MainDocumentPart.Document;

                // Delete CommentRangeStart for each
                // deleted comment in the main document.
                List<CommentRangeStart> commentRangeStartToDelete =
                    doc.Descendants<CommentRangeStart>().
                    Where(c => commentIds.Contains(c.Id.Value)).ToList();
                foreach (CommentRangeStart c in commentRangeStartToDelete)
                {
                    c.Remove();
                }

                // Delete CommentRangeEnd for each deleted comment in the main document.
                List<CommentRangeEnd> commentRangeEndToDelete =
                    doc.Descendants<CommentRangeEnd>().
                    Where(c => commentIds.Contains(c.Id.Value)).ToList();
                foreach (CommentRangeEnd c in commentRangeEndToDelete)
                {
                    c.Remove();
                }

                // Delete CommentReference for each deleted comment in the main document.
                List<CommentReference> commentRangeReferenceToDelete =
                    doc.Descendants<CommentReference>().
                    Where(c => commentIds.Contains(c.Id.Value)).ToList();
                foreach (CommentReference c in commentRangeReferenceToDelete)
                {
                    c.Remove();
                }

                // Save changes back to the MainDocumentPart part.
                doc.Save();
            }
        }
    }
}
