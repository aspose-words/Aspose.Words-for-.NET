// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class DeleteCommentsByAllOrASpecificAuthor : TestUtil
    {
        [Test]
        public void DeleteCommentsByAllOrASpecificAuthorFeature()
        {
            RemoveComments("");
        }

        private void RemoveComments(string authorName)
        {
            Document doc = new Document(MyDir + "Comments.docx");

            // Collect all comments in the document.
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);
            if (authorName == "")
            {
                // Remove all comments.
                comments.Clear();
            }
            else
            {
                // Look through all comments and remove those written by the authorName author.
                for (int i = comments.Count - 1; i >= 0; i--)
                {
                    Comment comment = (Comment)comments[i];
                    if (comment.Author == authorName)
                        comment.Remove();
                }
            }

            doc.Save(ArtifactsDir + "Remove comments - Aspose.Words.docx");
        }
    }
}
