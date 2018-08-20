// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.Collections.Generic;
using Aspose.Words;

namespace AddingComments
{
    public static class ReplyToCommentHelper
    {
        public static List<Comment> ExtractReplies(Document doc)
        {
            List<Comment> collectedComments = new List<Comment>();

            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            foreach (Comment comment in comments)
            {
                foreach (Comment reply in comment.Replies)
                {
                    collectedComments.Add(reply);
                }
            }

            return collectedComments;
        }

        public static List<Comment> ExtractReplies(Comment comment)
        {
            List<Comment> collectedComments = new List<Comment>();

            foreach (Comment reply in comment.Replies)
            {
                collectedComments.Add(reply);
            }

            return collectedComments;
        }

        public static void RemoveReplyAt(Comment comment, int replyIndex)
        {
            comment.RemoveReply(comment.Replies[replyIndex]);
        }

        public static void RemoveReplies(Document doc)
        {
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            foreach (Comment comment in comments)
            {
                if (comment.Ancestor != null)
                {
                    comment.Remove();
                }
            }
        }

        public static bool IsReply(Comment comment)
        {
            if (comment.Ancestor != null)
            {
                return true;
            }

            return false;
        }

        public static void MarkRepliesAsDone(Comment comment)
        {
            foreach (Comment reply in comment.Replies)
            {
                reply.Done = true;
            }
        }
    }
}
