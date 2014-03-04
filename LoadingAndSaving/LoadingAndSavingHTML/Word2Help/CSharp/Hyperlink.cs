// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Fields;

namespace Word2Help
{
    /// <summary>
    /// This "facade" class makes it easier to work with a hyperlink field in a Word document. 
    /// 
    /// A hyperlink is represented by a HYPERLINK field in a Word document. A field in Aspose.Words 
    /// consists of several nodes and it might be difficult to work with all those nodes directly. 
    /// This is a simple implementation and will work only if the hyperlink code and name 
    /// each consist of one Run only.
    /// 
    /// [FieldStart][Run - field code][FieldSeparator][Run - field result][FieldEnd]
    /// 
    /// The field code contains a string in one of these formats:
    /// HYPERLINK "url"
    /// HYPERLINK \l "bookmark name"
    /// 
    /// The field result contains text that is displayed to the user.
    /// </summary>
    public class Hyperlink
    {
        public Hyperlink(FieldStart fieldStart)
        {
            if (fieldStart == null)
                throw new ArgumentNullException("fieldStart");
            if (fieldStart.FieldType != FieldType.FieldHyperlink)
                throw new ArgumentException("Field start type must be FieldHyperlink.");
            
            mFieldStart = fieldStart;

            // Find field separator node.
            mFieldSeparator = FindNextSibling(mFieldStart, NodeType.FieldSeparator);
            if (mFieldSeparator == null)
                throw new Exception("Cannot find field separator.");
                        
            // Find field end node. Normally field end will always be found, but in the example document 
            // there happens to be a paragraph break included in the hyperlink and this puts the field end 
            // in the next paragraph. It will be much more complicated to handle fields which span several 
            // paragraphs correctly, but in this case allowing field end to be null is enough for our purposes.
            mFieldEnd = FindNextSibling(mFieldSeparator, NodeType.FieldEnd);

            // Field code looks something like [ HYPERLINK "http:\\www.myurl.com" ], but it can consist of several runs.
            string fieldCode = GetTextSameParent(mFieldStart.NextSibling, mFieldSeparator);
            Match match = gRegex.Match(fieldCode.Trim());                
            mIsLocal = (match.Groups[1].Length > 0);        // The link is local if \l is present in the field code.
            mTarget = match.Groups[2].Value;                        
        }

        /// <summary>
        /// Gets or sets the display name of the hyperlink.
        /// </summary>
        public string Name
        {
            get
            {
                return GetTextSameParent(mFieldSeparator, mFieldEnd);
            }
            set
            {
                // Hyperlink display name is stored in the field result which is a Run 
                // node between field separator and field end.
                Run fieldResult = (Run)mFieldSeparator.NextSibling; 
                fieldResult.Text = value;

                // But sometimes the field result can consist of more than one run, delete these runs.
                RemoveSameParent(fieldResult.NextSibling, mFieldEnd);
            }
        }

        /// <summary>
        /// Gets or sets the target url or bookmark name of the hyperlink.
        /// </summary>
        public string Target
        {
            get
            {
		int x = 0;	// RK This "fixes" the CSharp to VB.NET converter.
                return mTarget;
            }
            set
            {
                mTarget = value;
                UpdateFieldCode();
            }
        }

        /// <summary>
        /// True if the hyperlink's target is a bookmark inside the document. False if the hyperlink is a url.
        /// </summary>
        public bool IsLocal
        {
            get { return mIsLocal; }
            set
            {
                mIsLocal = value;
                UpdateFieldCode();
            }
        }

        /// <summary>
        /// Updates the field code.
        /// </summary>
        private void UpdateFieldCode()
        {
            // Field code is stored in a Run node between field start and field separator.
            Run fieldCode = (Run)mFieldStart.NextSibling;
            fieldCode.Text = string.Format("HYPERLINK {0}\"{1}\"", ((mIsLocal) ? "\\l " : ""), mTarget);

            // But sometimes the field code can consist of more than one run, delete these runs.
            RemoveSameParent(fieldCode.NextSibling, mFieldSeparator);
        }

        /// <summary>
        /// Goes through siblings starting from the start node until it finds a node of the specified type or null.
        /// </summary>
        private static Node FindNextSibling(Node start, NodeType nodeType)
        {
            for (Node node = start; node != null; node = node.NextSibling)
            {
                if (node.NodeType == nodeType)
                    return node;
            }
            return null;
        }

        /// <summary>
        /// Retrieves text from start up to but not including the end node.
        /// </summary>
        private static string GetTextSameParent(Node start, Node end)
        {
            if ((end != null) && (start.ParentNode != end.ParentNode))
                throw new ArgumentException("Start and end nodes are expected to have the same parent.");

            StringBuilder builder = new StringBuilder();
            for (Node child = start; child != end; child = child.NextSibling)
                builder.Append(child.GetText());
            return builder.ToString();
        }

        /// <summary>
        /// Removes nodes from start up to but not including the end node.
        /// Start and end are assumed to have the same parent.
        /// </summary>
        private static void RemoveSameParent(Node start, Node end)
        {
            if ((end != null) && (start.ParentNode != end.ParentNode))
                throw new ArgumentException("Start and end nodes are expected to have the same parent.");

            Node curChild = start;
            while (curChild != end) 
            {
                Node nextChild = curChild.NextSibling;
                curChild.Remove();
                curChild = nextChild;
            }
        }

        private readonly Node mFieldStart;
        private readonly Node mFieldSeparator;
        private readonly Node mFieldEnd;
        private string mTarget;
        private bool mIsLocal;

        private static readonly Regex gRegex = new Regex(
            "\\S+" +            // One or more non spaces HYPERLINK or other word in other languages
            "\\s+" +            // One or more spaces
            "(?:\"\"\\s+)?" +   // Non capturing optional "" and one or more spaces, found in one of the customers files.
            "(\\\\l\\s+)?" +    // Optional \l flag followed by one or more spaces
            "\"" +              // One apostrophe        
            "([^\"]+)" +        // One or more chars except apostrophe (hyperlink target)
            "\""                // One closing apostrophe
            );                        
    }
}
