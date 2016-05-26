// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

//ExStart
//ExFor:NodeList
//ExFor:FieldStart
//ExId:ReplaceHyperlinks
//ExSummary:Finds all hyperlinks in a Word document and changes their URL and display name.

using System;
using System.Text;
using System.Text.RegularExpressions;

using Aspose.Words;
using Aspose.Words.Fields;

using NUnit.Framework;

//ExSkip

namespace ApiExamples
{
    /// <summary>
    /// Shows how to replace hyperlinks in a Word document.
    /// </summary>
    [TestFixture] //ExSkip
    public class ExReplaceHyperlinks : ApiExampleBase
    {
        /// <summary>
        /// Finds all hyperlinks in a Word document and changes their URL and display name.
        /// </summary>
        [Test] //ExSkip
        public void ReplaceHyperlinks()
        {
            // Specify your document name here.
            Document doc = new Document(MyDir + "ReplaceHyperlinks.doc");

            // Hyperlinks in a Word documents are fields, select all field start nodes so we can find the hyperlinks.
            NodeList fieldStarts = doc.SelectNodes("//FieldStart");
            foreach (FieldStart fieldStart in fieldStarts)
            {
                if (fieldStart.FieldType.Equals(FieldType.FieldHyperlink))
                {
                    // The field is a hyperlink field, use the "facade" class to help to deal with the field.
                    Hyperlink hyperlink = new Hyperlink(fieldStart);

                    // Some hyperlinks can be local (links to bookmarks inside the document), ignore these.
                    if (hyperlink.IsLocal)
                        continue;

                    // The Hyperlink class allows to set the target URL and the display name 
                    // of the link easily by setting the properties.
                    hyperlink.Target = NewUrl;
                    hyperlink.Name = NewName;
                }
            }

            doc.Save(MyDir + @"\Artifacts\ReplaceHyperlinks.doc");
        }

        private const string NewUrl = @"http://www.aspose.com";
        private const string NewName = "Aspose - The .NET & Java Component Publisher";
    }


    /// <summary>
    /// This "facade" class makes it easier to work with a hyperlink field in a Word document. 
    /// 
    /// A hyperlink is represented by a HYPERLINK field in a Word document. A field in Aspose.Words 
    /// consists of several nodes and it might be difficult to work with all those nodes directly. 
    /// Note this is a simple implementation and will work only if the hyperlink code and name 
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
    internal class Hyperlink
    {
        internal Hyperlink(FieldStart fieldStart)
        {
            if (fieldStart == null)
                throw new ArgumentNullException("fieldStart");
            if (!fieldStart.FieldType.Equals(FieldType.FieldHyperlink))
                throw new ArgumentException("Field start type must be FieldHyperlink.");
            
            this.mFieldStart = fieldStart;

            // Find the field separator node.
            this.mFieldSeparator = FindNextSibling(this.mFieldStart, NodeType.FieldSeparator);
            if (this.mFieldSeparator == null)
                throw new InvalidOperationException("Cannot find field separator.");
            
            // Find the field end node. Normally field end will always be found, but in the example document 
            // there happens to be a paragraph break included in the hyperlink and this puts the field end 
            // in the next paragraph. It will be much more complicated to handle fields which span several 
            // paragraphs correctly, but in this case allowing field end to be null is enough for our purposes.
            this.mFieldEnd = FindNextSibling(this.mFieldSeparator, NodeType.FieldEnd);

            // Field code looks something like [ HYPERLINK "http:\\www.myurl.com" ], but it can consist of several runs.
            string fieldCode = GetTextSameParent(this.mFieldStart.NextSibling, this.mFieldSeparator);
            Match match = gRegex.Match(fieldCode.Trim());		
            this.mIsLocal = (match.Groups[1].Length > 0);	//The link is local if \l is present in the field code.
            this.mTarget = match.Groups[2].Value;			
        }

        /// <summary>
        /// Gets or sets the display name of the hyperlink.
        /// </summary>
        internal string Name
        {
            get
            {
                return GetTextSameParent(this.mFieldSeparator, this.mFieldEnd);
            }
            set
            {
                // Hyperlink display name is stored in the field result which is a Run 
                // node between field separator and field end.
                Run fieldResult = (Run)this.mFieldSeparator.NextSibling; 
                fieldResult.Text = value;

                // But sometimes the field result can consist of more than one run, delete these runs.
                RemoveSameParent(fieldResult.NextSibling, this.mFieldEnd);
            }
        }

        /// <summary>
        /// Gets or sets the target url or bookmark name of the hyperlink.
        /// </summary>
        internal string Target
        {
            get
            {
                string dummy = null;  // This is needed to fool the C# to VB.NET converter.
                return this.mTarget;
            }
            set
            {
                this.mTarget = value;
                this.UpdateFieldCode();
            }
        }

        /// <summary>
        /// True if the hyperlink's target is a bookmark inside the document. False if the hyperlink is a url.
        /// </summary>
        internal bool IsLocal
        {
            get
            {
                return this.mIsLocal;
            }
            set
            {
                this.mIsLocal = value;
                this.UpdateFieldCode();
            }
        }

        private void UpdateFieldCode()
        {
            // Field code is stored in a Run node between field start and field separator.
            Run fieldCode = (Run)this.mFieldStart.NextSibling;
            fieldCode.Text = string.Format("HYPERLINK {0}\"{1}\"", ((this.mIsLocal) ? "\\l " : ""), this.mTarget);

            // But sometimes the field code can consist of more than one run, delete these runs.
            RemoveSameParent(fieldCode.NextSibling, this.mFieldSeparator);
        }

        /// <summary>
        /// Goes through siblings starting from the start node until it finds a node of the specified type or null.
        /// </summary>
        private static Node FindNextSibling(Node startNode, NodeType nodeType)
        {
            for (Node node = startNode; node != null; node = node.NextSibling)
            {
                if (node.NodeType.Equals(nodeType))
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
        /// Start and end are assumed to have the same parent.
        /// </summary>
        private static void RemoveSameParent(Node startNode, Node endNode)
        {
            if ((endNode != null) && (startNode.ParentNode != endNode.ParentNode))
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

        /// <summary>
        /// RK I am notoriously bad at regexes. It seems I don't understand their way of thinking.
        /// </summary>
        private static readonly Regex gRegex = new Regex(
            "\\S+" +			// one or more non spaces HYPERLINK or other word in other languages
            "\\s+" +			// one or more spaces
            "(?:\"\"\\s+)?" +	// non capturing optional "" and one or more spaces, found in one of the customers files.
            "(\\\\l\\s+)?" +	// optional \l flag followed by one or more spaces
            "\"" +				// one apostrophe	
            "([^\"]+)" +		// one or more chars except apostrophe (hyperlink target)
            "\""				// one closing apostrophe
            );			
    }
}
//ExEnd