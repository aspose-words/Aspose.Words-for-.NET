// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using Aspose.Words;

namespace AddingComments
{
    public static class CommentsHelper
    {
        public static Comment AddComment(Document doc, string authorName, string initials, DateTime dateTime,
            string commentText)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);

            Comment comment = new Comment(doc, authorName, initials, dateTime);
            comment.SetText(commentText);

            builder.CurrentParagraph.AppendChild(comment);

            return comment;
        }

        public static List<Comment> ExtractComments(Document doc)
        {
            List<Comment> collectedComments = new List<Comment>();

            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            foreach (Comment comment in comments)
            {
                collectedComments.Add(comment);
            }

            return collectedComments;
        }
        
        public static List<Comment> ExtractComments(Document doc, string authorName)
        {
            List<Comment> collectedComments = new List<Comment>();

            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            foreach (Comment comment in comments)
            {
                if (comment.Author == authorName)
                    collectedComments.Add(comment);
            }

            return collectedComments;
        }

        public static void RemoveComments(Document doc)
        {
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            comments.Clear();
        }

        public static void RemoveComments(Document doc, string authorName)
        {
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            for (int i = comments.Count - 1; i >= 0; i--)
            {
                Comment comment = (Comment)comments[i];
                if (comment.Author == authorName)
                    comment.Remove();
            }
        }

        public static void MarkCommentsAsDone(Document doc)
        {
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            foreach (Comment comment in comments)
            {
                if (comment.Ancestor == null)
                {
                    comment.Done = true;
                }
            }
        }

        public static void MarkCommentAsDone(Comment comment)
        {
            comment.Done = true;
        }
    }
}
