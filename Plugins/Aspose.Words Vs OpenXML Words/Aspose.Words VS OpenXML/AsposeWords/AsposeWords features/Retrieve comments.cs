// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System;
using System.Collections;
using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class RetrieveComments : TestUtil
    {
        [Test]
        public void RetrieveCommentsFeature()
        {
            Document doc = new Document(MyDir + "Comments.docx");

            ArrayList collectedComments = new ArrayList();
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            // Look through all comments and gather information about them.
            foreach (Comment comment in comments)
                collectedComments.Add(comment.Author + " " + comment.DateTime + " " + comment.ToString(SaveFormat.Text));

            foreach (string collectedComment in collectedComments)
                Console.WriteLine(collectedComment);
        }
    }
}
