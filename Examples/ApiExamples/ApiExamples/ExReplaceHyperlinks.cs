// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

//ExStart
//ExFor:NodeList
//ExFor:FieldStart
//ExSummary:Shows how to find all hyperlinks in a Word document, and then change their URLs and display names.
using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture] //ExSkip
    public class ExReplaceHyperlinks : ApiExampleBase
    {
        [Test] //ExSkip
        public void Fields()
        {
            Document doc = new Document(MyDir + "Hyperlinks.docx");

            // Hyperlinks in a Word documents are fields. To begin looking for hyperlinks, we must first find all the fields.
            // Use the "SelectNodes" method to find all the fields in the document via an XPath.
            NodeList fieldStarts = doc.SelectNodes("//FieldStart");

            foreach (FieldStart fieldStart in fieldStarts.OfType<FieldStart>())
            {
                if (fieldStart.FieldType == FieldType.FieldHyperlink)
                {
                    Hyperlink hyperlink = new Hyperlink(fieldStart);

                    // Hyperlinks that link to bookmarks do not have URLs.
                    if (hyperlink.IsLocal)
                        continue;

                    // Give each URL hyperlink a new URL and name.
                    hyperlink.Target = NewUrl;
                    hyperlink.Name = NewName;
                }
            }

            doc.Save(ArtifactsDir + "ReplaceHyperlinks.Fields.docx");
        }

        private const string NewUrl = @"http://www.aspose.com";
        private const string NewName = "Aspose - The .NET & Java Component Publisher";
    }

    /// <summary>
    /// HYPERLINK fields contain and display hyperlinks in the document body. A field in Aspose.Words 
    /// consists of several nodes, and it might be difficult to work with all those nodes directly. 
    /// This implementation will work only if the hyperlink code and name each consist of only one Run node.
    ///
    /// The node structure for fields is as follows:
    /// 
    /// [FieldStart][Run - field code][FieldSeparator][Run - field result][FieldEnd]
    /// 
    /// Below are two example field codes of HYPERLINK fields:
    /// HYPERLINK "url"
    /// HYPERLINK \l "bookmark name"
    /// 
    /// A field's "Result" property contains text that the field displays in the document body to the user.
    /// </summary>
    internal class Hyperlink
    {
        internal Hyperlink(FieldStart fieldStart)
        {
            if (fieldStart == null)
                throw new ArgumentNullException("fieldStart");
            if (fieldStart.FieldType != FieldType.FieldHyperlink)
                throw new ArgumentException("Field start type must be FieldHyperlink.");

            mFieldStart = fieldStart;

            // Find the field separator node.
            mFieldSeparator = FindNextSibling(mFieldStart, NodeType.FieldSeparator);
            if (mFieldSeparator == null)
                throw new InvalidOperationException("Cannot find field separator.");

            // Normally, we can always find the field's end node, but the example document 
            // contains a paragraph break inside a hyperlink, which puts the field end 
            // in the next paragraph. It will be much more complicated to handle fields which span several 
            // paragraphs correctly. In this case allowing field end to be null is enough.
            mFieldEnd = FindNextSibling(mFieldSeparator, NodeType.FieldEnd);

            // Field code looks something like "HYPERLINK "http:\\www.myurl.com"", but it can consist of several runs.
            string fieldCode = GetTextSameParent(mFieldStart.NextSibling, mFieldSeparator);
            Match match = gRegex.Match(fieldCode.Trim());

            // The hyperlink is local if \l is present in the field code.
            mIsLocal = match.Groups[1].Length > 0; 
            mTarget = match.Groups[2].Value;
        }

        /// <summary>
        /// Gets or sets the display name of the hyperlink.
        /// </summary>
        internal string Name
        {
            get => GetTextSameParent(mFieldSeparator, mFieldEnd); 
            set
            {
                // Hyperlink display name is stored in the field result, which is a Run 
                // node between field separator and field end.
                Run fieldResult = (Run) mFieldSeparator.NextSibling;
                fieldResult.Text = value;

                // If the field result consists of more than one run, delete these runs.
                RemoveSameParent(fieldResult.NextSibling, mFieldEnd);
            }
        }

        /// <summary>
        /// Gets or sets the target URL or bookmark name of the hyperlink.
        /// </summary>
        internal string Target
        {
            get => mTarget;
            set
            {
                mTarget = value;
                UpdateFieldCode();
            }
        }

        /// <summary>
        /// True if the hyperlinks target is a bookmark inside the document. False if the hyperlink is a URL.
        /// </summary>
        internal bool IsLocal
        {
            get => mIsLocal; 
            set
            {
                mIsLocal = value;
                UpdateFieldCode();
            }
        }

        private void UpdateFieldCode()
        {
            // A field's field code is in a Run node between the field's start node and field separator.
            Run fieldCode = (Run) mFieldStart.NextSibling;
            fieldCode.Text = string.Format("HYPERLINK {0}\"{1}\"", ((mIsLocal) ? "\\l " : ""), mTarget);

            // If the field code consists of more than one run, delete these runs.
            RemoveSameParent(fieldCode.NextSibling, mFieldSeparator);
        }

        /// <summary>
        /// Goes through siblings starting from the start node until it finds a node of the specified type or null.
        /// </summary>
        private static Node FindNextSibling(Node startNode, NodeType nodeType)
        {
            for (Node node = startNode; node != null; node = node.NextSibling)
            {
                if (node.NodeType == nodeType)
                    return node;
            }

            return null;
        }

        /// <summary>
        /// Retrieves text from start up to but not including the end node.
        /// </summary>
        private static string GetTextSameParent(Node startNode, Node endNode)
        {
            if ((endNode != null) && (startNode.ParentNode != endNode.ParentNode))
                throw new ArgumentException("Start and end nodes are expected to have the same parent.");

            StringBuilder builder = new StringBuilder();
            for (Node child = startNode; !child.Equals(endNode); child = child.NextSibling)
                builder.Append(child.GetText());

            return builder.ToString();
        }

        /// <summary>
        /// Removes nodes from start up to but not including the end node.
        /// Assumes that the start and end nodes have the same parent.
        /// </summary>
        private static void RemoveSameParent(Node startNode, Node endNode)
        {
            if (endNode != null && startNode.ParentNode != endNode.ParentNode)
                throw new ArgumentException("Start and end nodes are expected to have the same parent.");

            Node curChild = startNode;
            while ((curChild != null) && (curChild != endNode))
            {
                Node nextChild = curChild.NextSibling;
                curChild.Remove();
                curChild = nextChild;
            }
        }

        private readonly Node mFieldStart;
        private readonly Node mFieldSeparator;
        private readonly Node mFieldEnd;
        private bool mIsLocal;
        private string mTarget;

        private static readonly Regex gRegex = new Regex(
            "\\S+" + // One or more non spaces HYPERLINK or other word in other languages.
            "\\s+" + // One or more spaces.
            "(?:\"\"\\s+)?" + // Non-capturing optional "" and one or more spaces.
            "(\\\\l\\s+)?" + // Optional \l flag followed by one or more spaces.
            "\"" + // One apostrophe.	
            "([^\"]+)" + // One or more characters, excluding the apostrophe (hyperlink target).
            "\"" // One closing apostrophe.
        );
    }
}
//ExEnd