// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class RemoveSpecificComments : TestUtil
    {
        [Test]
        public void RemoveCommentsAsposeWords()
        {
            //ExStart:RemoveCommentsAsposeWords
            //GistId:787486ce8310219ee50379944022f5db
            Document doc = new Document(MyDir + "Comments.docx");

            // Collect all comments in the document.
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            string authorName = string.Empty;
            if (authorName.Equals(string.Empty))
                // Remove all comments.
                comments.Clear();
            else
            {
                // Remove comments by author name.
                foreach (Comment comment in comments)
                    if (comment.Author.Equals(authorName))
                        comment.Remove();
            }

            doc.Save(ArtifactsDir + "Remove comments - Aspose.Words.docx");
            //ExEnd:RemoveCommentsAsposeWords
        }
    }
}
