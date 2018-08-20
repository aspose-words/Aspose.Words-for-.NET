// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using Aspose.Words;
using NUnit.Framework;

namespace AddingComments
{
    /// <summary>
    /// This project demonstrates how to create and work with comments <see cref="RunCommentExamples"/> and 
    /// replies to comments <see cref="RunReplyToCommentExamples"/> on the AW document.
    /// </summary>
    [TestFixture]
    public class Program
    {
        [Test]
        public static void Main()
        {
            RunCommentExamples(); 
            RunReplyToCommentExamples();
        }

        public static void RunCommentExamples()
        {
            // Create new test document for adding comments and replies.
            Document doc = new Document();

            // Add test comments to the document
            for (int i = 0; i <= 10; i++)
            {
                CommentsHelper.AddComment(doc, "Author " + i, "Initials " + i, DateTime.Now, "Comment text " + i);
            }

            Console.WriteLine("Comments are added!");

            // Extract the information about the comments of all the authors.
            foreach (Comment comment in CommentsHelper.ExtractComments(doc))
                Console.Write(comment.GetText());

            // Remove comments by the "Author 2" author.
            CommentsHelper.RemoveComments(doc, "Author 2");
            Console.WriteLine("Comments are removed!");

            // Extract the information about the comments of the "Author 1" author and mark as "Done"
            foreach (Comment comment in CommentsHelper.ExtractComments(doc, "Author 1"))
                CommentsHelper.MarkCommentAsDone(comment);

            // Mark all comments in the document as "Done"
            CommentsHelper.MarkCommentsAsDone(doc);
            Console.WriteLine("All comments marks as 'Done'");

            // Remove all comments.
            CommentsHelper.RemoveComments(doc);
            Console.WriteLine("All comments are  removed!");
        }

        public static void RunReplyToCommentExamples()
        {
            // Create new test document for adding comments and replies.
            Document doc = new Document();

            // Add comments with replies to the document
            for (int i = 0; i <= 10; i++)
            {
                Comment comment = CommentsHelper.AddComment(doc, "Author " + i, "Initials " + i, DateTime.Now, "Comment text " + i);

                for (int y = 0; y <= 10; y++)
                {
                    comment.AddReply("Reply author " + y, "Reply initials " + y, DateTime.Now, "Reply to comment " + y);
                }
            }

            Console.WriteLine("All comments and replies are added!");

            // Extract the information about  all the replies of all the comments in the document
            foreach (Comment reply in ReplyToCommentHelper.ExtractReplies(doc))
            {
                Console.Write(reply.Ancestor.GetText());
                Console.Write(reply.GetText());
            }

            // Remove reply to comment by index.
            foreach (Comment comment in CommentsHelper.ExtractComments(doc, "Author 1"))
                ReplyToCommentHelper.RemoveReplyAt(comment, 2);

            Console.WriteLine("Reply was removed!");

            // Extract the information about the replies from comment of the "Author 2" author and mark as "Done"
            foreach (Comment comment in CommentsHelper.ExtractComments(doc, "Author 2"))
                foreach (Comment reply in ReplyToCommentHelper.ExtractReplies(comment))
                reply.Done = true;

            // Mark replies of the "Author 1" comment author as "Done"
            foreach (Comment comment in CommentsHelper.ExtractComments(doc, "Author 3"))
                ReplyToCommentHelper.MarkRepliesAsDone(comment);

            Console.WriteLine("All replies marks as 'Done'");

            // Remove reply of the "Author 4" comment author by index
            foreach (Comment comment in CommentsHelper.ExtractComments(doc, "Author 4"))
                ReplyToCommentHelper.RemoveReplyAt(comment, 1);

            // Remove all replies of the "Author 4" comment author
            foreach (Comment comment in CommentsHelper.ExtractComments(doc, "Author 4"))
                comment.RemoveAllReplies();

            Console.WriteLine("Specific replies are removed!");

            // Check that comment is reply to
            foreach (Comment comment in CommentsHelper.ExtractComments(doc, "Author 5"))
                if (ReplyToCommentHelper.IsReply(comment))
                    Console.WriteLine(comment.GetText());

            // Remove all replies to comments from the document
            ReplyToCommentHelper.RemoveReplies(doc);

            Console.WriteLine("All replies are removed!");
        }
    }
}
