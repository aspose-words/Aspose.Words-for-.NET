// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Xml.XPath;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using NUnit.Framework;
#if NETSTANDARD2_0 || __MOBILE__
using SkiaSharp;
#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExNode : ApiExampleBase
    {
        [Test]
        public void CloneCompositeNode()
        {
            //ExStart
            //ExFor:Node
            //ExFor:Node.Clone
            //ExSummary:Shows how to clone composite nodes with and without their child nodes.
            // Create a new empty document
            Document doc = new Document();

            // Add some text to the first paragraph
            Paragraph para = doc.FirstSection.Body.FirstParagraph;
            para.AppendChild(new Run(doc, "Some text"));

            // Clone the paragraph and the child nodes
            Node cloneWithChildren = para.Clone(true);
            // Only clone the paragraph and no child nodes
            Node cloneWithoutChildren = para.Clone(false);
            //ExEnd

            Assert.IsTrue(((CompositeNode) cloneWithChildren).HasChildNodes);
            Assert.IsFalse(((CompositeNode) cloneWithoutChildren).HasChildNodes);
        }

        [Test]
        public void GetParentNode()
        {
            //ExStart
            //ExFor:Node.ParentNode
            //ExSummary:Shows how to access the parent node.
            // Create a new empty document. It has one section
            Document doc = new Document();

            // The section is the first child node of the document
            Node section = doc.FirstChild;

            // The section's parent node is the document
            Console.WriteLine("Section parent is the document: " + (doc == section.ParentNode));
            //ExEnd

            Assert.AreEqual(doc, section.ParentNode);
        }

        [Test]
        public void OwnerDocument()
        {
            //ExStart
            //ExFor:Node.Document
            //ExFor:Node.ParentNode
            //ExSummary:Shows that when you create any node, it requires a document that will own the node.
            // Open a file from disk
            Document doc = new Document();

            // Creating a new node of any type requires a document passed into the constructor
            Paragraph para = new Paragraph(doc);

            // The new paragraph node does not yet have a parent
            Console.WriteLine("Paragraph has no parent node: " + (para.ParentNode == null));

            // But the paragraph node knows its document
            Console.WriteLine("Both nodes' documents are the same: " + (para.Document == doc));

            // The fact that a node always belongs to a document allows us to access and modify 
            // properties that reference the document-wide data such as styles or lists
            para.ParagraphFormat.StyleName = "Heading 1";

            // Now add the paragraph to the main text of the first section
            doc.FirstSection.Body.AppendChild(para);

            // The paragraph node is now a child of the Body node
            Console.WriteLine("Paragraph has a parent node: " + (para.ParentNode != null));
            //ExEnd

            Assert.AreEqual(doc, para.Document);
            Assert.IsNotNull(para.ParentNode);
        }

        [Test]
        public void EnumerateChildNodes()
        {
            Document doc = new Document();
            //ExStart
            //ExFor:Node
            //ExFor:NodeType
            //ExFor:CompositeNode
            //ExFor:CompositeNode.GetChild
            //ExSummary:Shows how to extract a specific child node from a CompositeNode by using the GetChild method and passing the NodeType and index.
            Paragraph paragraph = (Paragraph) doc.GetChild(NodeType.Paragraph, 0, true);
            //ExEnd

            //ExStart
            //ExFor:CompositeNode.ChildNodes
            //ExFor:CompositeNode.GetEnumerator
            //ExSummary:Shows how to enumerate immediate children of a CompositeNode using the enumerator provided by the ChildNodes collection.
            NodeCollection children = paragraph.ChildNodes;
            foreach (Node child in children)
            {
                // Paragraph may contain children of various types such as runs, shapes and so on
                if (child.NodeType.Equals(NodeType.Run))
                {
                    // Say we found the node that we want, do something useful
                    Run run = (Run) child;
                    Console.WriteLine(run.Text);
                }
            }

            //ExEnd
        }

        [Test]
        public void IndexChildNodes()
        {
            Document doc = new Document();
            Paragraph paragraph = (Paragraph) doc.GetChild(NodeType.Paragraph, 0, true);

            //ExStart
            //ExFor:NodeCollection.Count
            //ExFor:NodeCollection.Item
            //ExSummary:Shows how to enumerate immediate children of a CompositeNode using indexed access.
            NodeCollection children = paragraph.ChildNodes;
            for (int i = 0; i < children.Count; i++)
            {
                Node child = children[i];

                // Paragraph may contain children of various types such as runs, shapes and so on
                if (child.NodeType.Equals(NodeType.Run))
                {
                    // Say we found the node that we want, do something useful
                    Run run = (Run) child;
                    Console.WriteLine(run.Text);
                }
            }

            //ExEnd
        }

        //ExStart
        //ExFor:Node.NextSibling
        //ExFor:CompositeNode.FirstChild
        //ExFor:Node.IsComposite
        //ExFor:CompositeNode.IsComposite
        //ExFor:Node.NodeTypeToString
        //ExSummary:Shows how to efficiently visit all direct and indirect children of a composite node.
        [Test] //ExSkip
        public void RecurseAllNodes()
        {
            // Open a document
            Document doc = new Document(MyDir + "Node.RecurseAllNodes.doc");

            // Invoke the recursive function that will walk the tree
            TraverseAllNodes(doc);
        }

        /// <summary>
        /// A simple function that will walk through all children of a specified node recursively 
        /// and print the type of each node to the screen.
        /// </summary>
        public void TraverseAllNodes(CompositeNode parentNode)
        {
            // This is the most efficient way to loop through immediate children of a node
            for (Node childNode = parentNode.FirstChild; childNode != null; childNode = childNode.NextSibling)
            {
                // Do some useful work
                Console.WriteLine(Node.NodeTypeToString(childNode.NodeType));

                // Recurse into the node if it is a composite node
                if (childNode.IsComposite)
                    TraverseAllNodes((CompositeNode) childNode);
            }
        }
        //ExEnd

        [Test]
        public void RemoveNodes()
        {
            Document doc = new Document();

            //ExStart
            //ExFor:Node
            //ExFor:Node.NodeType
            //ExFor:Node.Remove
            //ExSummary:Shows how to remove all nodes of a specific type from a composite node.
            // In this example we remove tables from a section body
            // Get the section that we want to work on
            Section section = doc.Sections[0];
            Body body = section.Body;

            // Select the first child node in the body
            Node curNode = body.FirstChild;

            while (curNode != null)
            {
                // Save the pointer to the next sibling node because if the current 
                // node is removed from the parent in the next step, we will have 
                // no way of finding the next node to continue the loop
                Node nextNode = curNode.NextSibling;

                // A section body can contain Paragraph and Table nodes
                // If the node is a Table, remove it from the parent
                if (curNode.NodeType.Equals(NodeType.Table))
                    curNode.Remove();

                // Continue going through child nodes until null (no more siblings) is reached
                curNode = nextNode;
            }

            //ExEnd
        }

        [Test]
        public void EnumNextSibling()
        {
            Document doc = new Document();

            //ExStart
            //ExFor:CompositeNode.FirstChild
            //ExFor:Node.NextSibling
            //ExFor:Node.NodeTypeToString
            //ExFor:Node.NodeType
            //ExSummary:Shows how to enumerate immediate child nodes of a composite node using NextSibling.
            // In this example we enumerate all paragraphs of a section body
            // Get the section that we want to work on
            Section section = doc.Sections[0];
            Body body = section.Body;

            // Loop starting from the first child until we reach null
            for (Node node = body.FirstChild; node != null; node = node.NextSibling)
            {
                // Output the types of the nodes that we come across
                Console.WriteLine(Node.NodeTypeToString(node.NodeType));
            }

            //ExEnd
        }

        [Test]
        public void TypedAccess()
        {
            Document doc = new Document();

            //ExStart
            //ExFor:Story.Tables
            //ExFor:Table.FirstRow
            //ExFor:Table.LastRow
            //ExFor:TableCollection
            //ExSummary:Shows how to use typed properties to access nodes of the document tree.
            // Quick typed access to the first child Section node of the Document
            Section section = doc.FirstSection;

            // Quick typed access to the Body child node of the Section
            Body body = section.Body;

            // Quick typed access to all Table child nodes contained in the Body
            TableCollection tables = body.Tables;

            foreach (Table table in tables.OfType<Table>())
            {
                // Quick typed access to the first row of the table
                table.FirstRow?.Remove();

                // Quick typed access to the last row of the table
                table.LastRow?.Remove();
            }
            //ExEnd
        }

        [Test]
        public void UpdateFieldsInRange()
        {
            Document doc = new Document();

            //ExStart
            //ExFor:Range.UpdateFields
            //ExSummary:Shows how to update document fields in the body of the first section only.
            doc.FirstSection.Body.Range.UpdateFields();
            //ExEnd
        }

        [Test]
        public void RemoveChild()
        {
            Document doc = new Document();

            //ExStart
            //ExFor:CompositeNode.LastChild
            //ExFor:Node.PreviousSibling
            //ExFor:CompositeNode.RemoveChild
            //ExSummary:Shows how to use of methods of Node and CompositeNode to remove a section before the last section in the document.
            // Document is a CompositeNode and LastChild returns the last child node in the Document node
            // Since the Document can contain only Section nodes, the last child is the last section
            Node lastSection = doc.LastChild;

            // Each node knows its next and previous sibling nodes
            // Previous sibling of a section is a section before the specified section
            // If the node is the first child, PreviousSibling will return null
            Node sectionBeforeLast = lastSection.PreviousSibling;

            if (sectionBeforeLast != null)
                doc.RemoveChild(sectionBeforeLast);
            //ExEnd
        }

        [Test]
        public void CompositeNode_SelectNodes()
        {
            //ExStart
            //ExFor:CompositeNode.SelectSingleNode
            //ExFor:CompositeNode.SelectNodes
            //ExFor:NodeList.GetEnumerator
            //ExFor:NodeList.ToArray
            //ExSummary:Shows how to select certain nodes by using an XPath expression.
            Document doc = new Document(MyDir + "Table.Document.doc");

            // This expression will extract all paragraph nodes which are descendants of any table node in the document
            // This will return any paragraphs which are in a table
            NodeList nodeList = doc.SelectNodes("//Table//Paragraph");

            // Iterate through the list with an enumerator and print the contents of every paragraph in each cell of the table
            int index = 0;
            using (IEnumerator<Node> e = nodeList.GetEnumerator())
            {
                while (e.MoveNext())
                {
                    Console.WriteLine($"Table paragraph index {index++}, contents: \"{e.Current.GetText().Trim()}\"");
                }
            }

            // This expression will select any paragraphs that are direct children of any body node in the document
            nodeList = doc.SelectNodes("//Body/Paragraph");

            // We can treat the list as an array too
            Assert.AreEqual(4, nodeList.ToArray().Length);

            // Use SelectSingleNode to select the first result of the same expression as above
            Node node = doc.SelectSingleNode("//Body/Paragraph");
            //ExEnd
        }

        [Test]
        public void TestNodeIsInsideField()
        {
            //ExStart:
            //ExFor:CompositeNode.SelectNodes
            //ExFor:CompositeNode.GetChild
            //ExSummary:Shows how to test if a node is inside a field by using an XPath expression.
            // Let's pick a document we know has some fields in
            Document doc = new Document(MyDir + "MailMerge.MergeImage.doc");

            // Let's say we want to check if the Run below is inside a field
            Run run = (Run) doc.GetChild(NodeType.Run, 5, true);

            // Evaluate the XPath expression. The resulting NodeList will contain all nodes found inside a field a field (between FieldStart 
            // and FieldEnd exclusive). There can however be FieldStart and FieldEnd nodes in the list if there are nested fields 
            // in the path. Currently does not find rare fields in which the FieldCode or FieldResult spans across multiple paragraphs
            NodeList resultList =
                doc.SelectNodes("//FieldStart/following-sibling::node()[following-sibling::FieldEnd]");

            // Check if the specified run is one of the nodes that are inside the field
            foreach (Node node in resultList)
            {
                if (node == run)
                {
                    Console.WriteLine("The node is found inside a field");
                    break;
                }
            }
            //ExEnd
        }

        [Test]
        public void CreateAndAddParagraphNode()
        {
            Document doc = new Document();

            Paragraph para = new Paragraph(doc);

            Section section = doc.LastSection;
            section.Body.AppendChild(para);
        }

        [Test]
        public void RemoveSmartTagsFromCompositeNode()
        {
            //ExStart
            //ExFor:CompositeNode.RemoveSmartTags
            //ExSummary:Removes all smart tags from descendant nodes of the composite node.
            Document doc = new Document(MyDir + "Document.doc");

            // Remove smart tags from the first paragraph in the document
            doc.FirstSection.Body.FirstParagraph.RemoveSmartTags();
            //ExEnd
        }

        [Test]
        public void GetIndexOfNode()
        {
            //ExStart
            //ExFor:CompositeNode.IndexOf
            //ExSummary:Shows how to get the index of a given child node from its parent.
            Document doc = new Document(MyDir + "Rendering.doc");

            // Get the body of the first section in the document
            Body body = doc.FirstSection.Body;
            // Retrieve the index of the last paragraph in the body
            int index = body.ChildNodes.IndexOf(body.LastParagraph);
            //ExEnd

            // Verify that the index is correct
            Assert.AreEqual(24, index);
        }

        [Test]
        public void GetNodeTypeEnums()
        {
            //ExStart
            //ExFor:Paragraph.NodeType
            //ExFor:Table.NodeType
            //ExFor:Node.NodeType
            //ExFor:Footnote.NodeType
            //ExFor:FormField.NodeType
            //ExFor:SmartTag.NodeType
            //ExFor:Cell.NodeType
            //ExFor:Row.NodeType
            //ExFor:Document.NodeType
            //ExFor:Comment.NodeType
            //ExFor:Run.NodeType
            //ExFor:Section.NodeType
            //ExFor:SpecialChar.NodeType
            //ExFor:Shape.NodeType
            //ExFor:FieldEnd.NodeType
            //ExFor:FieldSeparator.NodeType
            //ExFor:FieldStart.NodeType
            //ExFor:BookmarkStart.NodeType
            //ExFor:CommentRangeEnd.NodeType
            //ExFor:BuildingBlock.NodeType
            //ExFor:GlossaryDocument.NodeType
            //ExFor:BookmarkEnd.NodeType
            //ExFor:GroupShape.NodeType
            //ExFor:CommentRangeStart.NodeType
            //ExSummary:Shows how to retrieve the NodeType enumeration of nodes.
            Document doc = new Document(MyDir + "Document.doc");

            // Let's pick a node that we can't be quite sure of what type it is
            // In this case lets pick the first node of the first paragraph in the body of the document
            Node node = doc.FirstSection.Body.FirstParagraph.FirstChild;
            Console.WriteLine("NodeType of first child: " + Node.NodeTypeToString(node.NodeType));

            // This time let's pick a node that we know the type of
            // Create a new paragraph and a table node
            Paragraph para = new Paragraph(doc);
            Table table = new Table(doc);

            // Access to NodeType for typed nodes will always return their specific NodeType
            // i.e A paragraph node will always return NodeType.Paragraph, a table node will always return NodeType.Table
            Console.WriteLine("NodeType of Paragraph: " + Node.NodeTypeToString(para.NodeType));
            Console.WriteLine("NodeType of Table: " + Node.NodeTypeToString(table.NodeType));
            //ExEnd
        }

        [Test]
        public void ConvertNodeToHtmlWithDefaultOptions()
        {
            //ExStart
            //ExFor:Node.ToString(SaveFormat)
            //ExSummary:Exports the content of a node to String in HTML format using default options.
            Document doc = new Document(MyDir + "Document.doc");

            // Extract the last paragraph in the document to convert to HTML
            Node node = doc.LastSection.Body.LastParagraph;

            // When ToString is called using the SaveFormat overload then conversion is executed using default save options
            // When saving to HTML using default options the following settings are set:
            //   ExportImagesAsBase64 = true
            //   CssStyleSheetType = CssStyleSheetType.Inline
            //   ExportFontResources = false
            string nodeAsHtml = node.ToString(SaveFormat.Html);
            //ExEnd

            Assert.AreEqual(
                "<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:12pt\"><span style=\"font-family:'Times New Roman'\">Hello World!</span></p>",
                nodeAsHtml);
        }

        [Test]
        public void ConvertNodeToHtmlWithSaveOptions()
        {
            //ExStart
            //ExFor:Node.ToString(SaveOptions)
            //ExSummary:Exports the content of a node to String in HTML format using custom specified options.
            Document doc = new Document(MyDir + "Document.doc");

            // Extract the last paragraph in the document to convert to HTML
            Node node = doc.LastSection.Body.LastParagraph;

            // Create an instance of HtmlSaveOptions and set a few options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportHeadersFootersMode = ExportHeadersFootersMode.PerSection,
                ExportRelativeFontSize = true
            };

            // Convert the document to HTML and return as a String. Pass the instance of HtmlSaveOptions to
            // to use the specified options during the conversion
            string nodeAsHtml = node.ToString(saveOptions);
            //ExEnd

            Assert.AreEqual(
                "<p style=\"margin-top:0pt; margin-bottom:0pt\"><span style=\"font-family:'Times New Roman'\">Hello World!</span></p>",
                nodeAsHtml);
        }

        [Test]
        public void TypedNodeCollectionToArray()
        {
            Document doc = new Document();

            //ExStart
            //ExFor:ParagraphCollection.ToArray
            //ExSummary:Demonstrates typed implementations of ToArray on classes derived from NodeCollection.
            // You can use ToArray to return a typed array of nodes
            Paragraph[] paras = doc.FirstSection.Body.Paragraphs.ToArray();
            //ExEnd

            Assert.That(paras.Length, Is.GreaterThan(0));
        }

        [Test]
        public void NodeEnumerationHotRemove()
        {
            //ExStart
            //ExFor:ParagraphCollection.ToArray
            //ExSummary:Demonstrates how to use "hot remove" to remove a node during enumeration.
            DocumentBuilder builder = new DocumentBuilder();
            builder.Writeln("The first paragraph");
            builder.Writeln("The second paragraph");
            builder.Writeln("The third paragraph");
            builder.Writeln("The fourth paragraph");

            // Hot remove allows a node to be removed from a live collection and have the enumeration continue
            foreach (Paragraph para in builder.Document.FirstSection.Body.GetChildNodes(NodeType.Paragraph, true)
                .OfType<Paragraph>())
            {
                if (para.Range.Text.Contains("third"))
                {
                    // Enumeration will continue even after this node is removed
                    para.Remove();
                }
            }
            //ExEnd
        }

        [Test]
        public void EnumerationHotRemoveLimitations()
        {
            //ExStart
            //ExFor:ParagraphCollection.ToArray
            //ExSummary:Demonstrates an example breakage of the node collection enumerator.
            DocumentBuilder builder = new DocumentBuilder();
            builder.Writeln("The first paragraph");
            builder.Writeln("The second paragraph");
            builder.Writeln("The third paragraph");
            builder.Writeln("The fourth paragraph");

            // This causes unexpected behavior, the fourth paragraph in the collection is not visited
            foreach (Paragraph para in builder.Document.FirstSection.Body.GetChildNodes(NodeType.Paragraph, true)
                .OfType<Paragraph>())
            {
                if (para.Range.Text.Contains("third"))
                {
                    para.PreviousSibling.Remove();
                    para.Remove();
                }
            }
            //ExEnd
        }

        [Test]
        public void CompositeNodeChildren()
        {
            //ExStart
            //ExFor:CompositeNode.Count
            //ExFor:CompositeNode.GetChildNodes(NodeType[], Boolean)
            //ExFor:CompositeNode.InsertAfter(Node, Node)
            //ExFor:CompositeNode.InsertBefore(Node, Node)
            //ExFor:CompositeNode.PrependChild(Node) 
            //ExFor:Paragraph.GetText
            //ExSummary:Shows how to add, update and delete child nodes from within a CompositeNode.
            Document doc = new Document();

            // An empty document has one paragraph by default
            Assert.AreEqual(1, doc.FirstSection.Body.Paragraphs.Count);

            // A paragraph is a composite node because it can contain runs, which are another type of node
            Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
            Run paragraphText = new Run(doc, "Initial text. ");
            paragraph.AppendChild(paragraphText);

            // We will place these 3 children into the main text of our paragraph
            Run run1 = new Run(doc, "Run 1. ");
            Run run2 = new Run(doc, "Run 2. ");
            Run run3 = new Run(doc, "Run 3. ");

            // We initialized them but not in our paragraph yet
            Assert.AreEqual("Initial text. " + (char) 12, paragraph.GetText());

            // Insert run2 before initial paragraph text. This will be at the start of the paragraph
            paragraph.InsertBefore(run2, paragraphText);

            // Insert run3 after initial paragraph text. This will be at the end of the paragraph
            paragraph.InsertAfter(run3, paragraphText);

            // Insert run1 before every other child node. run2 was the start of the paragraph, now it will be run1
            paragraph.PrependChild(run1);

            Assert.AreEqual("Run 1. Run 2. Initial text. Run 3. " + (char) 12, paragraph.GetText());
            Assert.AreEqual(4, paragraph.GetChildNodes(NodeType.Any, true).Count);

            // Access the child node collection and update/delete children
            ((Run) paragraph.GetChildNodes(NodeType.Run, true)[1]).Text = "Updated run 2. ";
            paragraph.GetChildNodes(NodeType.Run, true).Remove(paragraphText);

            Assert.AreEqual("Run 1. Updated run 2. Run 3. " + (char) 12, paragraph.GetText());
            Assert.AreEqual(3, paragraph.GetChildNodes(NodeType.Any, true).Count);
            //ExEnd
        }

        //ExStart
        //ExFor:CompositeNode.CreateNavigator
        //ExSummary:Shows how to create an XPathNavigator and use it to traverse and read nodes.
        [Test] //ExSkip
        public void NodeXPathNavigator()
        {
            // Create a blank document
            Document doc = new Document();

            // A document is a composite node so we can make a navigator straight away
            XPathNavigator navigator = doc.CreateNavigator();

            // Our root is the document node with 1 child, which is the first section
            if (navigator != null)
            {
                Assert.AreEqual("Document", navigator.Name);
                Assert.AreEqual(false, navigator.MoveToNext());
                Assert.AreEqual(1, navigator.SelectChildren(XPathNodeType.All).Count);

                // The document tree has the document, first section, body and first paragraph as nodes, with each being an only child of the previous
                // We can add a few more to give the tree some branches for the navigator to traverse
                DocumentBuilder docBuilder = new DocumentBuilder(doc);
                docBuilder.Write("Section 1, Paragraph 1. ");
                docBuilder.InsertParagraph();
                docBuilder.Write("Section 1, Paragraph 2. ");
                doc.AppendChild(new Section(doc));
                docBuilder.MoveToSection(1);
                docBuilder.Write("Section 2, Paragraph 1. ");

                // Use our navigator to print a map of all the nodes in the document to the console
                StringBuilder stringBuilder = new StringBuilder();
                MapDocument(navigator, stringBuilder, 0);
                Console.Write(stringBuilder.ToString());
            }
        }

        /// <summary>
        /// This will traverse all children of a composite node and map the structure in the style of a directory tree.
        /// Amount of space indentation indicates depth relative to initial node. Only runs will have their values printed.
        /// </summary>
        private static void MapDocument(XPathNavigator navigator, StringBuilder stringBuilder, int depth)
        {
            do
            {
                stringBuilder.Append(' ', depth);
                stringBuilder.Append(navigator.Name + ": ");

                if (navigator.Name == "Run")
                {
                    stringBuilder.Append(navigator.Value);
                }

                stringBuilder.Append('\n');

                if (navigator.HasChildren)
                {
                    navigator.MoveToFirstChild();
                    MapDocument(navigator, stringBuilder, depth + 1);
                    navigator.MoveToParent();
                }
            } while (navigator.MoveToNext());
        }
        //ExEnd

        //ExStart
        //ExFor:NodeChangingAction
        //ExFor:NodeChangingArgs.Action
        //ExFor:NodeChangingArgs.NewParent
        //ExFor:NodeChangingArgs.OldParent
        //ExSummary:Shows how to use a NodeChangingCallback to monitor changes to the document tree as it is edited.
        [Test] //ExSkip
        public void NodeChangingCallback()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set the NodeChangingCallback attribute to a custom printer 
            doc.NodeChangingCallback = new NodeChangingPrinter();

            // All node additions and removals will be printed to the console as we edit the document
            builder.Writeln("Hello world!");
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndTable();

            #if NETFRAMEWORK
            builder.InsertImage(Image.FromFile(ImageDir + "Aspose.Words.gif"));
            #else
            using (SKBitmap image = SKBitmap.Decode(ImageDir + "Aspose.Words.gif"))
                builder.InsertImage(image);
            #endif

            builder.CurrentParagraph.ParentNode.RemoveAllChildren();
        }

        /// <summary>
        /// Prints all inserted/removed nodes as well as their parent nodes.
        /// </summary>
        private class NodeChangingPrinter : INodeChangingCallback
        {
            void INodeChangingCallback.NodeInserting(NodeChangingArgs args)
            {
                Assert.AreEqual(NodeChangingAction.Insert, args.Action);
                Assert.AreEqual(null, args.OldParent);
            }

            void INodeChangingCallback.NodeInserted(NodeChangingArgs args)
            {
                Assert.AreEqual(NodeChangingAction.Insert, args.Action);
                Assert.NotNull(args.NewParent);

                Console.WriteLine($"Inserted node:");
                Console.WriteLine($"\tType:\t{args.Node.NodeType}");

                if (args.Node.GetText().Trim() != "")
                {
                    Console.WriteLine($"\tText:\t\"{args.Node.GetText().Trim()}\"");
                }

                Console.WriteLine($"\tHash:\t{args.Node.GetHashCode()}");
                Console.WriteLine($"\tParent:\t{args.NewParent.NodeType} ({args.NewParent.GetHashCode()})");
            }

            void INodeChangingCallback.NodeRemoving(NodeChangingArgs args)
            {
                Assert.AreEqual(NodeChangingAction.Remove, args.Action);
            }

            void INodeChangingCallback.NodeRemoved(NodeChangingArgs args)
            {
                Assert.AreEqual(NodeChangingAction.Remove, args.Action);
                Assert.Null(args.NewParent);

                Console.WriteLine($"Removed node: {args.Node.NodeType} ({args.Node.GetHashCode()})");
            }
        }
        //ExEnd

        [Test]
        public void NodeCollection()
        {
            //ExStart
            //ExFor:NodeCollection.Contains(Node)
            //ExFor:NodeCollection.Insert(Int32,Node)
            //ExFor:NodeCollection.Remove(Node)
            //ExSummary:Shows how to work with a NodeCollection.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The normal way to insert Runs into a document is to add text using a DocumentBuilder
            builder.Write("Run 1. ");
            builder.Write("Run 2. ");

            // Every .Write() invocation creates a new Run, which is added to the parent Paragraph's RunCollection
            RunCollection runs = doc.FirstSection.Body.FirstParagraph.Runs;
            Assert.AreEqual(2, runs.Count);

            // We can insert a node into the RunCollection manually to achieve the same effect
            Run newRun = new Run(doc, "Run 3. ");
            runs.Insert(3, newRun);

            Assert.True(runs.Contains(newRun));
            Assert.AreEqual("Run 1. Run 2. Run 3.", doc.GetText().Trim());

            // Text can also be deleted from the document by accessing individual Runs via the RunCollection and editing or removing them
            Run run = runs[1];
            runs.Remove(run);
            Assert.AreEqual("Run 1. Run 3.", doc.GetText().Trim());

            Assert.NotNull(run);
            Assert.False(runs.Contains(run));
            //ExEnd
        }

        [Test]
        public void NodeList()
        {
            //ExStart
            //ExFor:NodeList.Count
            //ExFor:NodeList.Item(System.Int32)
            //ExSummary:Shows how to use XPaths to navigate a NodeList.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert some nodes with a DocumentBuilder
            builder.Writeln("Hello world!");

            builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndTable();

            #if NETFRAMEWORK
            builder.InsertImage(Image.FromFile(ImageDir + "Aspose.Words.gif"));
            #else
            using (SKBitmap image = SKBitmap.Decode(ImageDir + "Aspose.Words.gif"))
                builder.InsertImage(image);
            #endif

            // Get all run nodes, of which we put 3 in the entire document
            NodeList nodeList = doc.SelectNodes("//Run");
            Assert.AreEqual(3, nodeList.Count);

            // Using a double forward slash, select all Run nodes that are indirect descendants of a Table node,
            // which would in this case be the runs inside the two cells we inserted
            nodeList = doc.SelectNodes("//Table//Run");
            Assert.AreEqual(2, nodeList.Count);

            // Single forward slashes specify direct descendant relationships,
            // of which we skipped quite a few by using double slashes
            Assert.AreEqual(doc.SelectNodes("//Table//Run"), doc.SelectNodes("//Table/Row/Cell/Paragraph/Run"));

            // We can access the actual nodes via a NodeList too
            nodeList = doc.SelectNodes("//Shape");
            Assert.AreEqual(1, nodeList.Count);
            Shape shape = (Shape)nodeList[0];
            Assert.True(shape.HasImage);
            //ExEnd
        }
    }
}