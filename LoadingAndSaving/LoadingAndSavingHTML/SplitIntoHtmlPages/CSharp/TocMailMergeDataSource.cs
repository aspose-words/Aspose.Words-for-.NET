//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.Collections;
using Aspose.Words.MailMerging;

namespace SplitIntoHtmlPagesExample
{
    /// <summary>
    /// A custom data source for Aspose.Words mail merge.
    /// Returns topic objects.
    /// </summary>
    internal class TocMailMergeDataSource : IMailMergeDataSource
    {
        internal TocMailMergeDataSource(ArrayList topics)
        {
            mTopics = topics;
            // Initialize to BOF.
            mIndex = -1;
        }

        public bool MoveNext()
        {
            if (mIndex < mTopics.Count - 1)
            {
                mIndex++;
                return true;
            }
            else
            {
                // Reached EOF, return false.
                return false;
            }
        }

        public bool GetValue(string fieldName, out object fieldValue)
        {
            if (fieldName == "TocEntry")
            {
                // The template document is supposed to have only one field called "TocEntry".
                fieldValue = mTopics[mIndex];
                return true;
            }
            else
            {
                fieldValue = null;
                return false;
            }
        }

        public string TableName
        {
            // The template document is supposed to have only one merge region called "TOC".
            get { return "TOC"; }
        }

        public IMailMergeDataSource GetChildDataSource(string tableName)
        {
            return null;
        }

        private readonly ArrayList mTopics;
        private int mIndex;
    }
}