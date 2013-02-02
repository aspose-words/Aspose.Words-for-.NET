//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
namespace SplitIntoHtmlPages
{
    /// <summary>
    /// A simple class to hold a topic title and HTML file name together.
    /// </summary>
    internal class Topic
    {
        internal Topic(string title, string fileName)
        {
            mTitle = title;
            mFileName = fileName;
        }

        internal string Title
        {
            get { return mTitle; }
        }

        internal string FileName
        {
            get { return mFileName; }
        }

        private readonly string mTitle;
        private readonly string mFileName;
    }
}
