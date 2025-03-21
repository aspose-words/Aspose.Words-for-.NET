﻿// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace ApiExamples
{
    /// <summary>
    /// Shows how to rename merge fields in a Word document.
    /// </summary>
    [TestFixture]
    public class ExRenameMergeFields : ApiExampleBase
    {
        /// <summary>
        /// Finds all merge fields in a Word document and changes their names.
        /// </summary>
        [Test]
        public void Rename()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Dear ");
            builder.InsertField("MERGEFIELD  FirstName ");
            builder.Write(" ");
            builder.InsertField("MERGEFIELD  LastName ");
            builder.Writeln(",");
            builder.InsertField("MERGEFIELD  CustomGreeting ");

            // Select all field start nodes so we can find the MERGEFIELDs.
            NodeCollection fieldStarts = doc.GetChildNodes(NodeType.FieldStart, true);
            foreach (FieldStart fieldStart in fieldStarts.OfType<FieldStart>())
            {
                if (fieldStart.FieldType == FieldType.FieldMergeField)
                {
                    MergeField mergeField = new MergeField(fieldStart);
                    mergeField.Name = mergeField.Name + "_Renamed";
                }
            }

            doc.Save(ArtifactsDir + "RenameMergeFields.Rename.docx");
        }
    }

    /// <summary>
    /// Represents a facade object for a merge field in a Microsoft Word document.
    /// </summary>
    internal class MergeField
    {
        internal MergeField(FieldStart fieldStart)
        {
            if (fieldStart.FieldType != FieldType.FieldMergeField)
                throw new ArgumentException("Field start type must be FieldMergeField.");

            mFieldStart = fieldStart;

            // Find the field separator node.
            mFieldSeparator = FindNextSibling(mFieldStart, NodeType.FieldSeparator);
            if (mFieldSeparator == null)
                throw new InvalidOperationException("Cannot find field separator.");

            // Find the field end node. Normally field end will always be found, but in the example document 
            // there happens to be a paragraph break included in the hyperlink and this puts the field end 
            // in the next paragraph. It will be much more complicated to handle fields which span several 
            // paragraphs correctly, but in this case allowing field end to be null is enough for our purposes.
            mFieldEnd = FindNextSibling(mFieldSeparator, NodeType.FieldEnd);
        }

        /// <summary>
        /// Gets or sets the name of the merge field.
        /// </summary>
        internal string Name
        {
            get
            {
                return GetTextSameParent(mFieldSeparator.NextSibling, mFieldEnd).Trim('«', '»');
            }
            set
            {
                // Merge field name is stored in the field result which is a Run 
                // node between field separator and field end.
                Run fieldResult = (Run) mFieldSeparator.NextSibling;
                fieldResult.Text = $"«{value}»";

                // But sometimes the field result can consist of more than one run, delete these runs.
                RemoveSameParent(fieldResult.NextSibling, mFieldEnd);

                UpdateFieldCode(value);
            }
        }

        private void UpdateFieldCode(string fieldName)
        {
            // Field code is stored in a Run node between field start and field separator.
            Run fieldCode = (Run) mFieldStart.NextSibling;
            Match match = gRegex.Match(fieldCode.Text);

            string newFieldCode = $" {match.Groups["start"].Value}{fieldName} ";
            fieldCode.Text = newFieldCode;

            // But sometimes the field code can consist of more than one run, delete these runs.
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
            if (endNode != null && startNode.ParentNode != endNode.ParentNode)
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

        private static readonly Regex gRegex = new Regex(@"\s*(?<start>MERGEFIELD\s|)(\s|)(?<name>\S+)\s+");
    }
}