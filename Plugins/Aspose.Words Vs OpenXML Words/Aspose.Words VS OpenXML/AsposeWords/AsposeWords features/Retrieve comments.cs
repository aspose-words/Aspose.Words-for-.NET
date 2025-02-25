// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
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

            List<string> collectedComments = new();
            NodeCollection comments = doc.GetChildNodes(NodeType.Comment, true);

            // Look through all comments and gather information about them.
            foreach (Comment comment in comments)
                collectedComments.Add(comment.Author + " " + comment.DateTime + " " + comment.ToString(SaveFormat.Text));

            foreach (string collectedComment in collectedComments)
                Console.WriteLine(collectedComment);
        }
    }
}
