//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Data;
using System.IO;
using System.Reflection;
using System.Collections;

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

namespace ExtractContent
{
    class Program
    {
        private static string mDataDir;

        public static void Main(string[] args)
        {
            // Sample infrastructure.
            string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar;
            mDataDir = new Uri(new Uri(exeDir), @"../../Data/").LocalPath;

            // Call methods to test extraction of different types from the document.
            ExtractContentBetweenParagraphs();
            ExtractContentBetweenBlockLevelNodes();
            ExtractContentBetweenParagraphStyles();
            ExtractContentBetweenRuns();
            ExtractContentUsingField();
            ExtractContentBetweenBookmark();
            ExtractContentBetweenCommentRange();
        }

        public static void ExtractContentBetweenParagraphs()
        {
            //ExStart
            //ExId:ExtractBetweenNodes_BetweenParagraphs
            //ExSummary:Shows how to extract the content between specific paragraphs using the ExtractContent method above.
            // Load in the document
            Document doc = new Document(mDataDir + "TestFile.doc");

            // Gather the nodes. The GetChild method uses 0-based index
            Paragraph startPara = (Paragraph)doc.FirstSection.GetChild(NodeType.Paragraph, 6, true);
            Paragraph endPara = (Paragraph)doc.FirstSection.GetChild(NodeType.Paragraph, 10, true);
            // Extract the content between these nodes in the document. Include these markers in the extraction.
            ArrayList extractedNodes = ExtractContent(startPara, endPara, true);

            // Insert the content into a new separate document and save it to disk.
            Document dstDoc = GenerateDocument(doc, extractedNodes);
            dstDoc.Save(mDataDir + "TestFile.Paragraphs Out.doc");
            //ExEnd
        }

        public static void ExtractContentBetweenBlockLevelNodes()
        {
            //ExStart
            //ExId:ExtractBetweenNodes_BetweenNodes
            //ExSummary:Shows how to extract the content between a paragraph and table using the ExtractContent method.
            // Load in the document
            Document doc = new Document(mDataDir + "TestFile.doc");

            Paragraph startPara = (Paragraph)doc.LastSection.GetChild(NodeType.Paragraph, 2, true);
            Table endTable = (Table)doc.LastSection.GetChild(NodeType.Table, 0, true);

            // Extract the content between these nodes in the document. Include these markers in the extraction.
            ArrayList extractedNodes = ExtractContent(startPara, endTable, true);

            // Lets reverse the array to make inserting the content back into the document easier.
            extractedNodes.Reverse();

            while (extractedNodes.Count > 0)
            {
                // Insert the last node from the reversed list 
                endTable.ParentNode.InsertAfter((Node)extractedNodes[0], endTable);
                // Remove this node from the list after insertion.
                extractedNodes.RemoveAt(0);
            }

            // Save the generated document to disk.
            doc.Save(mDataDir + "TestFile.DuplicatedContent Out.doc");
            //ExEnd
        }

        public static void ExtractContentBetweenParagraphStyles()
        {
            //ExStart
            //ExId:ExtractBetweenNodes_BetweenStyles
            //ExSummary:Shows how to extract content between paragraphs with specific styles using the ExtractContent method.
            // Load in the document
            Document doc = new Document(mDataDir + "TestFile.doc");

            // Gather a list of the paragraphs using the respective heading styles.
            ArrayList parasStyleHeading1 = ParagraphsByStyleName(doc, "Heading 1");
            ArrayList parasStyleHeading3 = ParagraphsByStyleName(doc, "Heading 3");

            // Use the first instance of the paragraphs with those styles.
            Node startPara1 = (Node)parasStyleHeading1[0];
            Node endPara1 = (Node)parasStyleHeading3[0];

            // Extract the content between these nodes in the document. Don't include these markers in the extraction.
            ArrayList extractedNodes = ExtractContent(startPara1, endPara1, false);

            // Insert the content into a new separate document and save it to disk.
            Document dstDoc = GenerateDocument(doc, extractedNodes);
            dstDoc.Save(mDataDir + "TestFile.Styles Out.doc");
            //ExEnd
        }

        public static void ExtractContentBetweenRuns()
        {
            //ExStart
            //ExId:ExtractBetweenNodes_BetweenRuns
            //ExSummary:Shows how to extract content between specific runs of the same paragraph using the ExtractContent method.
            // Load in the document
            Document doc = new Document(mDataDir + "TestFile.doc");

            // Retrieve a paragraph from the first section.
            Paragraph para = (Paragraph)doc.GetChild(NodeType.Paragraph, 7, true);
            
            // Use some runs for extraction.
            Run startRun = para.Runs[1];
            Run endRun = para.Runs[4];

            // Extract the content between these nodes in the document. Include these markers in the extraction.
            ArrayList extractedNodes = ExtractContent(startRun, endRun, true);

            // Get the node from the list. There should only be one paragraph returned in the list.
            Node node = (Node)extractedNodes[0];
            // Print the text of this node to the console.
            Console.WriteLine(node.ToString(SaveFormat.Text));
            //ExEnd
        }

        public static void ExtractContentUsingField()
        {
            //ExStart
            //ExFor:DocumentBuilder.MoveToMergeField(String, Boolean, Boolean)
            //ExId:ExtractBetweenNodes_UsingField
            //ExSummary:Shows how to extract content between a specific field and paragraph in the document using the ExtractContent method.
            // Load in the document
            Document doc = new Document(mDataDir + "TestFile.doc");

            // Use a document builder to retrieve the field start of a merge field.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Pass the first boolean parameter to get the DocumentBuilder to move to the FieldStart of the field.
            // We could also get FieldStarts of a field using GetChildNode method as in the other examples.
            builder.MoveToMergeField("Fullname", false, false);

            // The builder cursor should be positioned at the start of the field.
            FieldStart startField = (FieldStart)builder.CurrentNode;
            Paragraph endPara = (Paragraph)doc.FirstSection.GetChild(NodeType.Paragraph, 5, true);

            // Extract the content between these nodes in the document. Don't include these markers in the extraction.
            ArrayList extractedNodes = ExtractContent(startField, endPara, false);

            // Insert the content into a new separate document and save it to disk.
            Document dstDoc = GenerateDocument(doc, extractedNodes);
            dstDoc.Save(mDataDir + "TestFile.Fields Out.pdf");
            //ExEnd
        }

        public static void ExtractContentBetweenBookmark()
        {
            //ExStart
            //ExId:ExtractBetweenNodes_BetweenBookmark
            //ExSummary:Shows how to extract the content referenced a bookmark using the ExtractContent method.
            // Load in the document
            Document doc = new Document(mDataDir + "TestFile.doc");

            // Retrieve the bookmark from the document.
            Aspose.Words.Bookmark bookmark = doc.Range.Bookmarks["Bookmark1"];

            // We use the BookmarkStart and BookmarkEnd nodes as markers.
            BookmarkStart bookmarkStart = bookmark.BookmarkStart;
            BookmarkEnd bookmarkEnd = bookmark.BookmarkEnd;

            // Firstly extract the content between these nodes including the bookmark. 
            ArrayList extractedNodesInclusive = ExtractContent(bookmarkStart, bookmarkEnd, true);
            Document dstDoc = GenerateDocument(doc, extractedNodesInclusive);
            dstDoc.Save(mDataDir + "TestFile.BookmarkInclusive Out.doc");

            // Secondly extract the content between these nodes this time without including the bookmark.
            ArrayList extractedNodesExclusive = ExtractContent(bookmarkStart, bookmarkEnd, false);
            dstDoc = GenerateDocument(doc, extractedNodesExclusive);
            dstDoc.Save(mDataDir + "TestFile.BookmarkExclusive Out.doc");
            //ExEnd
        }

        public static void ExtractContentBetweenCommentRange()
        {
            //ExStart
            //ExId:ExtractBetweenNodes_BetweenComment
            //ExSummary:Shows how to extract content referenced by a comment using the ExtractContent method.
            // Load in the document
            Document doc = new Document(mDataDir + "TestFile.doc");

            // This is a quick way of getting both comment nodes.
            // Your code should have a proper method of retrieving each corresponding start and end node.
            CommentRangeStart commentStart = (CommentRangeStart)doc.GetChild(NodeType.CommentRangeStart, 0, true);
            CommentRangeEnd commentEnd = (CommentRangeEnd)doc.GetChild(NodeType.CommentRangeEnd, 0, true);

            // Firstly extract the content between these nodes including the comment as well. 
            ArrayList extractedNodesInclusive = ExtractContent(commentStart, commentEnd, true);
            Document dstDoc = GenerateDocument(doc, extractedNodesInclusive);
            dstDoc.Save(mDataDir + "TestFile.CommentInclusive Out.doc");

            // Secondly extract the content between these nodes without the comment.
            ArrayList extractedNodesExclusive = ExtractContent(commentStart, commentEnd, false);
            dstDoc = GenerateDocument(doc, extractedNodesExclusive);
            dstDoc.Save(mDataDir + "TestFile.CommentExclusive Out.doc");
            //ExEnd
        }

        //ExStart
        //ExId:ExtractBetweenNodes_ExtractContent
        //ExSummary:This is a method which extracts blocks of content from a document between specified nodes.
        /// <summary>
        /// Extracts a range of nodes from a document found between specified markers and returns a copy of those nodes. Content can be extracted
        /// between inline nodes, block level nodes, and also special nodes such as Comment or Boomarks. Any combination of different marker types can used.
        /// </summary>
        /// <param name="startNode">The node which defines where to start the extraction from the document. This node can be block or inline level of a body.</param>
        /// <param name="endNode">The node which defines where to stop the extraction from the document. This node can be block or inline level of body.</param>
        /// <param name="isInclusive">Should the marker nodes be included.</returns>
        public static ArrayList ExtractContent(Node startNode, Node endNode, bool isInclusive)
        {
            // First check that the nodes passed to this method are valid for use.
            VerifyParameterNodes(startNode, endNode);

            // Create a list to store the extracted nodes.
            ArrayList nodes = new ArrayList();

            // Keep a record of the original nodes passed to this method so we can split marker nodes if needed.
            Node originalStartNode = startNode;
            Node originalEndNode = endNode;

            // Extract content based on block level nodes (paragraphs and tables). Traverse through parent nodes to find them.
            // We will split the content of first and last nodes depending if the marker nodes are inline
            while (startNode.ParentNode.NodeType != NodeType.Body)
                startNode = startNode.ParentNode;

            while (endNode.ParentNode.NodeType != NodeType.Body)
                endNode = endNode.ParentNode;

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
                CompositeNode cloneNode = (CompositeNode)currNode.Clone(true);
                isEndingNode = currNode.Equals(endNode);

                if(isStartingNode || isEndingNode)
                {
                    // We need to process each marker separately so pass it off to a separate method instead.
                    if (isStartingNode)
                    {
                        ProcessMarker(cloneNode, nodes, originalStartNode, isInclusive, isStartingNode, isEndingNode);
                        isStartingNode = false;
                    }

                    // Conditional needs to be separate as the block level start and end markers maybe the same node.
                    if (isEndingNode)
                    {
                        ProcessMarker(cloneNode, nodes, originalEndNode, isInclusive, isStartingNode, isEndingNode);
                        isExtracting = false;
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

            // Return the nodes between the node markers.
            return nodes;
        }
        //ExEnd

        //ExStart
        //ExId:ExtractBetweenNodes_Helpers
        //ExSummary:The helper methods used by the ExtractContent method.
        /// <summary>
        /// Checks the input parameters are correct and can be used. Throws an exception if there is any problem.
        /// </summary>
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
                if (startSection.Body.IndexOf(startNode) > endSection.Body.IndexOf(endNode))
                    throw new ArgumentException("The end node must be after the start node in the body");
            }
            else if (startIndex > endIndex)
                throw new ArgumentException("The section of end node must be after the section start node");
        }

        /// <summary>
        /// Checks if a node passed is an inline node.
        /// </summary>
        private static bool IsInline(Node node)
        {
            // Test if the node is desendant of a Paragraph or Table node and also is not a paragraph or a table a paragraph inside a comment class which is decesant of a pararaph is possible.
            return ((node.GetAncestor(NodeType.Paragraph) != null || node.GetAncestor(NodeType.Table) != null) && !(node.NodeType == NodeType.Paragraph || node.NodeType == NodeType.Table));
        }

        /// <summary>
        /// Removes the content before or after the marker in the cloned node depending on the type of marker.
        /// </summary>
        private static void ProcessMarker(CompositeNode cloneNode, ArrayList nodes, Node node, bool isInclusive, bool isStartMarker, bool isEndMarker)
        {
            // If we are dealing with a block level node just see if it should be included and add it to the list.
            if(!IsInline(node))
            {
                // Don't add the node twice if the markers are the same node
                if(!(isStartMarker && isEndMarker))
                {
                    if (isInclusive)
                        nodes.Add(cloneNode);
                }
                return;
            }

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

            // If either marker is part of a comment then to include the comment itself we need to move the pointer forward to the Comment
            // node found after the CommentRangeEnd node.
            if (node.NodeType == NodeType.CommentRangeEnd)
            {
                while (node.NextSibling != null && node.NodeType != NodeType.Comment)
                    node = node.NextSibling;

            }

            // Find the corresponding node in our cloned node by index and return it.
            // If the start and end node are the same some child nodes might already have been removed. Subtract the
            // difference to get the right index.
            int indexDiff = node.ParentNode.ChildNodes.Count - cloneNode.ChildNodes.Count;

            // Child node count identical.
            if (indexDiff == 0)
                node = cloneNode.ChildNodes[node.ParentNode.IndexOf(node)];
            else
                node = cloneNode.ChildNodes[node.ParentNode.IndexOf(node) - indexDiff];

            // Remove the nodes up to/from the marker.
            bool isSkip = false;
            bool isProcessing = true;
            bool isRemoving = isStartMarker;
            Node nextNode = cloneNode.FirstChild;

            while (isProcessing && nextNode != null)
            {
                Node currentNode = nextNode;
                isSkip = false;

                if (currentNode.Equals(node))
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

            // After processing the composite node may become empty. If it has don't include it.
            if (!(isStartMarker && isEndMarker))
            {
                if (cloneNode.HasChildNodes)
                    nodes.Add(cloneNode);
            }

        }
        //ExEnd

        //ExStart
        //ExId:ExtractBetweenNodes_GenerateDocument
        //ExSummary:This method takes a list of nodes and inserts them into a new document.
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
        //ExEnd

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

    }
}
