﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;

using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class Common
    {
        // ExStart:CommonExtractContent
        public static ArrayList ExtractContent(Node startNode, Node endNode, bool isInclusive)
        {
            // First check that the nodes passed to this method are valid for use.
            VerifyParameterNodes(startNode, endNode);

            // Create a list to store the extracted nodes.
            ArrayList nodes = new ArrayList();

            // If either marker is part of a comment then to include the comment itself we need to move the pointer
            // forward to the Comment Node found after the CommentRangeEnd node.
            if (endNode.NodeType == NodeType.CommentRangeEnd && isInclusive)
            {
                Node node = FindNextNode(NodeType.Comment, endNode.NextSibling);
                if (node != null)
                    endNode = node;
            }

            // Keep a record of the original nodes passed to this method so we can split marker nodes if needed.
            Node originalStartNode = startNode;
            Node originalEndNode = endNode;

            // Extract content based on block level nodes (paragraphs and tables). Traverse through parent nodes to find them.
            // We will split the content of first and last nodes depending if the marker nodes are inline
            startNode = GetAncestorInBody(startNode);
            endNode = GetAncestorInBody(endNode);

            bool isExtracting = true;
            bool isStartingNode = true;
            bool isEndingNode = false;
            // The current node we are extracting from the document.
            Node currNode = startNode;

            // Begin extracting content. Process all block level nodes and specifically split the first and last nodes when needed so paragraph formatting is retained.
            // Method is little more complex than a regular extractor as we need to factor in extracting using inline nodes, fields, bookmarks etc as to make it really useful.
            while (isExtracting)
            {
                // Clone the current node and its children to obtain a copy.
                Node cloneNode = currNode.Clone(true);
                isEndingNode = currNode.Equals(endNode);

                if (isStartingNode || isEndingNode)
                {
                    // We need to process each marker separately so pass it off to a separate method instead.
                    // End should be processed at first to keep node indexes.
                    if (isEndingNode)
                    {
                        // !isStartingNode: don't add the node twice if the markers are the same node.
                        ProcessMarker(cloneNode, nodes, originalEndNode, currNode, isInclusive,
                            false, !isStartingNode, false);
                        isExtracting = false;
                    }

                    // Conditional needs to be separate as the block level start and end markers maybe the same node.
                    if (isStartingNode)
                    {
                        ProcessMarker(cloneNode, nodes, originalStartNode, currNode, isInclusive,
                            true, true, false);
                        isStartingNode = false;
                    }
                }
                else
                    // Node is not a start or end marker, simply add the copy to the list.
                    nodes.Add(cloneNode);

                // Move to the next node and extract it. If next node is null that means the rest of the content is found in a different section.
                if (currNode.NextSibling == null && isExtracting)
                {
                    // Move to the next section.
                    Section nextSection = (Section)currNode.GetAncestor(NodeType.Section).NextSibling;
                    currNode = nextSection.Body.FirstChild;
                }
                else
                {
                    // Move to the next node in the body.
                    currNode = currNode.NextSibling;
                }
            }

            // For compatibility with mode with inline bookmarks, add the next paragraph (empty).
            if (isInclusive && originalEndNode == endNode && !originalEndNode.IsComposite)
                IncludeNextParagraph(endNode, nodes);

            // Return the nodes between the node markers.
            return nodes;
        }

        // ExEnd:CommonExtractContent
        public static ArrayList ParagraphsByStyleName(Document doc, string styleName)
        {
            // Create an array to collect paragraphs of the specified style.
            ArrayList paragraphsWithStyle = new ArrayList();
            // Get all paragraphs from the document.
            NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
            // Look through all paragraphs to find those with the specified style.
            foreach (Paragraph paragraph in paragraphs)
            {
                if (paragraph.ParagraphFormat.Style.Name == styleName)
                    paragraphsWithStyle.Add(paragraph);
            }
            return paragraphsWithStyle;
        }

        // ExStart:CommonGenerateDocument
        public static Document GenerateDocument(Document srcDoc, ArrayList nodes)
        {
            // Create a blank document.
            Document dstDoc = new Document();
            // Remove the first paragraph from the empty document.
            dstDoc.FirstSection.Body.RemoveAllChildren();

            // Import each node from the list into the new document. Keep the original formatting of the node.
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

            foreach (Node node in nodes)
            {
                Node importNode = importer.ImportNode(node, true);
                dstDoc.FirstSection.Body.AppendChild(importNode);
            }

            // Return the generated document.
            return dstDoc;
        }

        // ExEnd:CommonGenerateDocument  
        // ExStart:CommonExtractContentHelperMethods
        private static void VerifyParameterNodes(Node startNode, Node endNode)
        {
            // The order in which these checks are done is important.
            if (startNode == null)
                throw new ArgumentException("Start node cannot be null");
            if (endNode == null)
                throw new ArgumentException("End node cannot be null");

            if (!startNode.Document.Equals(endNode.Document))
                throw new ArgumentException("Start node and end node must belong to the same document");

            if (startNode.GetAncestor(NodeType.Body) == null || endNode.GetAncestor(NodeType.Body) == null)
                throw new ArgumentException("Start node and end node must be a child or descendant of a body");

            // Check the end node is after the start node in the DOM tree
            // First check if they are in different sections, then if they're not check their position in the body of the same section they are in.
            Section startSection = (Section)startNode.GetAncestor(NodeType.Section);
            Section endSection = (Section)endNode.GetAncestor(NodeType.Section);

            int startIndex = startSection.ParentNode.IndexOf(startSection);
            int endIndex = endSection.ParentNode.IndexOf(endSection);

            if (startIndex == endIndex)
            {
                if (startSection.Body.IndexOf(GetAncestorInBody(startNode)) >
                        endSection.Body.IndexOf(GetAncestorInBody(endNode)))
                    throw new ArgumentException("The end node must be after the start node in the body");
            }
            else if (startIndex > endIndex)
                throw new ArgumentException("The section of end node must be after the section start node");
        }

        private static Node FindNextNode(NodeType nodeType, Node fromNode)
        {
            if (fromNode == null || fromNode.NodeType == nodeType)
                return fromNode;

            if (fromNode.IsComposite)
            {
                Node node = FindNextNode(nodeType, ((CompositeNode)fromNode).FirstChild);
                if (node != null)
                    return node;
            }

            return FindNextNode(nodeType, fromNode.NextSibling);
        }

        private static bool IsInline(Node node)
        {
            // Test if the node is desendant of a Paragraph or Table node and also is not a paragraph or a table a paragraph inside a comment class which is decesant of a pararaph is possible.
            return ((node.GetAncestor(NodeType.Paragraph) != null || node.GetAncestor(NodeType.Table) != null) && !(node.NodeType == NodeType.Paragraph || node.NodeType == NodeType.Table));
        }

        private static void ProcessMarker(Node cloneNode, ArrayList nodes, Node node, Node blockLevelAncestor,
            bool isInclusive, bool isStartMarker, bool canAdd, bool forceAdd)
        {
            // If we are dealing with a block level node just see if it should be included and add it to the list.
            if (node == blockLevelAncestor)
            {
                if (canAdd && isInclusive)
                    nodes.Add(cloneNode);
                return;
            }
            // cloneNode is a clone of blockLevelNode. If node != blockLevelNode, blockLevelAncestor is ancestor of node
            // that means it is a composite node.
            Debug.Assert(cloneNode.IsComposite);

            // If a marker is a FieldStart node check if it's to be included or not.
            // We assume for simplicity that the FieldStart and FieldEnd appear in the same paragraph.
            if (node.NodeType == NodeType.FieldStart)
            {
                // If the marker is a start node and is not be included then skip to the end of the field.
                // If the marker is an end node and it is to be included then move to the end field so the field will not be removed.
                if ((isStartMarker && !isInclusive) || (!isStartMarker && isInclusive))
                {
                    while (node.NextSibling != null && node.NodeType != NodeType.FieldEnd)
                        node = node.NextSibling;
                }
            }

            // Support a case if marker node is on the third level of document body or lower.
            List<Node> nodeBranch = FillSelfAndParents(node, blockLevelAncestor);

            // Process the corresponding node in our cloned node by index.
            Node currentCloneNode = cloneNode;
            for (int i = nodeBranch.Count - 1; i >= 0; i--)
            {
                Node currentNode = nodeBranch[i];
                int nodeIndex = currentNode.ParentNode.IndexOf(currentNode);
                currentCloneNode = ((CompositeNode)currentCloneNode).ChildNodes[nodeIndex];

                RemoveNodesOutsideOfRange(currentCloneNode, isInclusive || (i > 0), isStartMarker);
            }

            // After processing the composite node may become empty. If it has don't include it.
            if (canAdd &&
                (forceAdd || ((CompositeNode)cloneNode).HasChildNodes))
                nodes.Add(cloneNode);
        }

        private static void RemoveNodesOutsideOfRange(Node markerNode, bool isInclusive, bool isStartMarker)
        {
            // Remove the nodes up to/from the marker.
            bool isSkip = false;
            bool isProcessing = true;
            bool isRemoving = isStartMarker;
            Node nextNode = markerNode.ParentNode.FirstChild;

            while (isProcessing && nextNode != null)
            {
                Node currentNode = nextNode;
                isSkip = false;

                if (currentNode.Equals(markerNode))
                {
                    if (isStartMarker)
                    {
                        isProcessing = false;
                        if (isInclusive)
                            isRemoving = false;
                    }
                    else
                    {
                        isRemoving = true;
                        if (isInclusive)
                            isSkip = true;
                    }
                }

                nextNode = nextNode.NextSibling;
                if (isRemoving && !isSkip)
                    currentNode.Remove();
            }
        }

        private static List<Node> FillSelfAndParents(Node node, Node tillNode)
        {
            List<Node> list = new List<Node>();
            Node currentNode = node;

            while (currentNode != tillNode)
            {
                list.Add(currentNode);
                currentNode = currentNode.ParentNode;
            }

            return list;
        }

        private static void IncludeNextParagraph(Node node, ArrayList nodes)
        {
            Paragraph paragraph = (Paragraph)FindNextNode(NodeType.Paragraph, node.NextSibling);
            if (paragraph != null)
            {
                // Move to first child to include paragraph without contents.
                Node markerNode = paragraph.HasChildNodes ? paragraph.FirstChild : paragraph;
                Node rootNode = GetAncestorInBody(paragraph);

                ProcessMarker(rootNode.Clone(true), nodes, markerNode, rootNode,
                    markerNode == paragraph, false, true, true);
            }
        }

        private static Node GetAncestorInBody(Node startNode)
        {
            while (startNode.ParentNode.NodeType != NodeType.Body)
                startNode = startNode.ParentNode;
            return startNode;
        }
        // ExEnd:CommonExtractContentHelperMethods
    }
}
