﻿// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.XPath;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using NUnit.Framework;
#if NET5_0_OR_GREATER || __MOBILE__
using SkiaSharp;
#else
using System.Drawing;
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
            //ExSummary:Shows how to clone a composite node.
            Document doc = new Document();
            Paragraph para = doc.FirstSection.Body.FirstParagraph;
            para.AppendChild(new Run(doc, "Hello world!"));

            // Below are two ways of cloning a composite node.
            // 1 -  Create a clone of a node, and create a clone of each of its child nodes as well.
            Node cloneWithChildren = para.Clone(true);

            Assert.That(((CompositeNode)cloneWithChildren).HasChildNodes, Is.True);
            Assert.That(cloneWithChildren.GetText().Trim(), Is.EqualTo("Hello world!"));

            // 2 -  Create a clone of a node just by itself without any children.
            Node cloneWithoutChildren = para.Clone(false);

            Assert.That(((CompositeNode)cloneWithoutChildren).HasChildNodes, Is.False);
            Assert.That(cloneWithoutChildren.GetText().Trim(), Is.EqualTo(string.Empty));
            //ExEnd
        }

        [Test]
        public void GetParentNode()
        {
            //ExStart
            //ExFor:Node.ParentNode
            //ExSummary:Shows how to access a node's parent node.
            Document doc = new Document();
            Paragraph para = doc.FirstSection.Body.FirstParagraph;

            // Append a child Run node to the document's first paragraph.
            Run run = new Run(doc, "Hello world!");
            para.AppendChild(run);

            // The paragraph is the parent node of the run node. We can trace this lineage
            // all the way to the document node, which is the root of the document's node tree.
            Assert.That(run.ParentNode, Is.EqualTo(para));
            Assert.That(para.ParentNode, Is.EqualTo(doc.FirstSection.Body));
            Assert.That(doc.FirstSection.Body.ParentNode, Is.EqualTo(doc.FirstSection));
            Assert.That(doc.FirstSection.ParentNode, Is.EqualTo(doc));
            //ExEnd
        }

        [Test]
        public void OwnerDocument()
        {
            //ExStart
            //ExFor:Node.Document
            //ExFor:Node.ParentNode
            //ExSummary:Shows how to create a node and set its owning document.
            Document doc = new Document();
            Paragraph para = new Paragraph(doc);
            para.AppendChild(new Run(doc, "Hello world!"));

            // We have not yet appended this paragraph as a child to any composite node.
            Assert.That(para.ParentNode, Is.Null);

            // If a node is an appropriate child node type of another composite node,
            // we can attach it as a child only if both nodes have the same owner document.
            // The owner document is the document we passed to the node's constructor.
            // We have not attached this paragraph to the document, so the document does not contain its text.
            Assert.That(doc, Is.EqualTo(para.Document));
            Assert.That(doc.GetText().Trim(), Is.EqualTo(string.Empty));

            // Since the document owns this paragraph, we can apply one of its styles to the paragraph's contents.
            para.ParagraphFormat.Style = doc.Styles["Heading 1"];

            // Add this node to the document, and then verify its contents.
            doc.FirstSection.Body.AppendChild(para);

            Assert.That(para.ParentNode, Is.EqualTo(doc.FirstSection.Body));
            Assert.That(doc.GetText().Trim(), Is.EqualTo("Hello world!"));
            //ExEnd

            Assert.That(para.Document, Is.EqualTo(doc));
            Assert.That(para.ParentNode, Is.Not.Null);
        }

        [Test]
        public void ChildNodesEnumerate()
        {
            //ExStart
            //ExFor:Node
            //ExFor:Node.CustomNodeId
            //ExFor:NodeType
            //ExFor:CompositeNode
            //ExFor:CompositeNode.GetChild
            //ExFor:CompositeNode.GetChildNodes(NodeType, bool)
            //ExFor:NodeCollection.Count
            //ExFor:NodeCollection.Item
            //ExSummary:Shows how to traverse through a composite node's collection of child nodes.
            Document doc = new Document();

            // Add two runs and one shape as child nodes to the first paragraph of this document.
            Paragraph paragraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
            paragraph.AppendChild(new Run(doc, "Hello world! "));

            Shape shape = new Shape(doc, ShapeType.Rectangle);
            shape.Width = 200;
            shape.Height = 200;
            // Note that the 'CustomNodeId' is not saved to an output file and exists only during the node lifetime.
            shape.CustomNodeId = 100;
            shape.WrapType = WrapType.Inline;
            paragraph.AppendChild(shape);

            paragraph.AppendChild(new Run(doc, "Hello again!"));

            // Iterate through the paragraph's collection of immediate children,
            // and print any runs or shapes that we find within.
            NodeCollection children = paragraph.GetChildNodes(NodeType.Any, false);

            Assert.That(paragraph.GetChildNodes(NodeType.Any, false).Count, Is.EqualTo(3));

            foreach (Node child in children)
                switch (child.NodeType)
                {
                    case NodeType.Run:
                        Console.WriteLine("Run contents:");
                        Console.WriteLine($"\t\"{child.GetText().Trim()}\"");
                        break;
                    case NodeType.Shape:
                        Shape childShape = (Shape)child;
                        Console.WriteLine("Shape:");
                        Console.WriteLine($"\t{childShape.ShapeType}, {childShape.Width}x{childShape.Height}");
                        Assert.That(shape.CustomNodeId, Is.EqualTo(100)); //ExSkip
                        break;
                }
            //ExEnd

            Assert.That(paragraph.GetChild(NodeType.Run, 0, true).NodeType, Is.EqualTo(NodeType.Run));
            Assert.That(doc.GetText().Trim(), Is.EqualTo("Hello world! Hello again!"));
        }

        //ExStart
        //ExFor:Node.NextSibling
        //ExFor:CompositeNode.FirstChild
        //ExFor:Node.IsComposite
        //ExFor:CompositeNode.IsComposite
        //ExFor:Node.NodeTypeToString
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
        //ExSummary:Shows how to traverse a composite node's tree of child nodes.
        [Test] //ExSkip
        public void RecurseChildren()
        {
            Document doc = new Document(MyDir + "Paragraphs.docx");

            // Any node that can contain child nodes, such as the document itself, is composite.
            Assert.That(doc.IsComposite, Is.True);

            // Invoke the recursive function that will go through and print all the child nodes of a composite node.
            TraverseAllNodes(doc, 0);
        }

        /// <summary>
        /// Recursively traverses a node tree while printing the type of each node
        /// with an indent depending on depth as well as the contents of all inline nodes.
        /// </summary>
        public void TraverseAllNodes(CompositeNode parentNode, int depth)
        {
            for (Node childNode = parentNode.FirstChild; childNode != null; childNode = childNode.NextSibling)
            {
                Console.Write($"{new string('\t', depth)}{Node.NodeTypeToString(childNode.NodeType)}");

                // Recurse into the node if it is a composite node. Otherwise, print its contents if it is an inline node.
                if (childNode.IsComposite)
                {
                    Console.WriteLine();
                    TraverseAllNodes((CompositeNode)childNode, depth + 1);
                }
                else if (childNode is Inline)
                {
                    Console.WriteLine($" - \"{childNode.GetText().Trim()}\"");
                }
                else
                {
                    Console.WriteLine();
                }
            }
        }
        //ExEnd

        [Test]
        public void RemoveNodes()
        {

            //ExStart
            //ExFor:Node
            //ExFor:Node.NodeType
            //ExFor:Node.Remove
            //ExSummary:Shows how to remove all child nodes of a specific type from a composite node.
            Document doc = new Document(MyDir + "Tables.docx");

            Assert.That(doc.GetChildNodes(NodeType.Table, true).Count, Is.EqualTo(2));

            Node curNode = doc.FirstSection.Body.FirstChild;

            while (curNode != null)
            {
                // Save the next sibling node as a variable in case we want to move to it after deleting this node.
                Node nextNode = curNode.NextSibling;

                // A section body can contain Paragraph and Table nodes.
                // If the node is a Table, remove it from the parent.
                if (curNode.NodeType == NodeType.Table)
                    curNode.Remove();

                curNode = nextNode;
            }

            Assert.That(doc.GetChildNodes(NodeType.Table, true).Count, Is.EqualTo(0));
            //ExEnd
        }

        [Test]
        public void EnumNextSibling()
        {
            //ExStart
            //ExFor:CompositeNode.FirstChild
            //ExFor:Node.NextSibling
            //ExFor:Node.NodeTypeToString
            //ExFor:Node.NodeType
            //ExSummary:Shows how to use a node's NextSibling property to enumerate through its immediate children.
            Document doc = new Document(MyDir + "Paragraphs.docx");

            for (Node node = doc.FirstSection.Body.FirstChild; node != null; node = node.NextSibling)
            {
                Console.WriteLine();
                Console.WriteLine($"Node type: {Node.NodeTypeToString(node.NodeType)}");

                string contents = node.GetText().Trim();
                Console.WriteLine(contents == string.Empty ? "This node contains no text" : $"Contents: \"{node.GetText().Trim()}\"");
            }
            //ExEnd
        }

        [Test]
        public void TypedAccess()
        {

            //ExStart
            //ExFor:Story.Tables
            //ExFor:Table.FirstRow
            //ExFor:Table.LastRow
            //ExFor:TableCollection
            //ExSummary:Shows how to remove the first and last rows of all tables in a document.
            Document doc = new Document(MyDir + "Tables.docx");

            TableCollection tables = doc.FirstSection.Body.Tables;

            Assert.That(tables[0].Rows.Count, Is.EqualTo(5));
            Assert.That(tables[1].Rows.Count, Is.EqualTo(4));

            foreach (Table table in tables.OfType<Table>())
            {
                table.FirstRow?.Remove();
                table.LastRow?.Remove();
            }

            Assert.That(tables[0].Rows.Count, Is.EqualTo(3));
            Assert.That(tables[1].Rows.Count, Is.EqualTo(2));
            //ExEnd
        }

        [Test]
        public void RemoveChild()
        {
            //ExStart
            //ExFor:CompositeNode.LastChild
            //ExFor:Node.PreviousSibling
            //ExFor:CompositeNode.RemoveChild``1(``0)
            //ExSummary:Shows how to use of methods of Node and CompositeNode to remove a section before the last section in the document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Writeln("Section 1 text.");
            builder.InsertBreak(BreakType.SectionBreakContinuous);
            builder.Writeln("Section 2 text.");

            // Both sections are siblings of each other.
            Section lastSection = (Section)doc.LastChild;
            Section firstSection = (Section)lastSection.PreviousSibling;

            // Remove a section based on its sibling relationship with another section.
            if (lastSection.PreviousSibling != null)
                doc.RemoveChild(firstSection);

            // The section we removed was the first one, leaving the document with only the second.
            Assert.That(doc.GetText().Trim(), Is.EqualTo("Section 2 text."));
            //ExEnd
        }

        [Test]
        public void SelectCompositeNodes()
        {
            //ExStart
            //ExFor:CompositeNode.SelectSingleNode
            //ExFor:CompositeNode.SelectNodes
            //ExFor:NodeList.GetEnumerator
            //ExFor:NodeList.ToArray
            //ExSummary:Shows how to select certain nodes by using an XPath expression.
            Document doc = new Document(MyDir + "Tables.docx");

            // This expression will extract all paragraph nodes,
            // which are descendants of any table node in the document.
            NodeList nodeList = doc.SelectNodes("//Table//Paragraph");

            // Iterate through the list with an enumerator and print the contents of every paragraph in each cell of the table.
            int index = 0;

            using (IEnumerator<Node> e = nodeList.GetEnumerator())
                while (e.MoveNext())
                    Console.WriteLine($"Table paragraph index {index++}, contents: \"{e.Current.GetText().Trim()}\"");

            // This expression will select any paragraphs that are direct children of any Body node in the document.
            nodeList = doc.SelectNodes("//Body/Paragraph");

            // We can treat the list as an array.
            Assert.That(nodeList.ToArray().Length, Is.EqualTo(4));

            // Use SelectSingleNode to select the first result of the same expression as above.
            Node node = doc.SelectSingleNode("//Body/Paragraph");

            Assert.That(node.GetType(), Is.EqualTo(typeof(Paragraph)));
            //ExEnd
        }

        [Test]
        public void TestNodeIsInsideField()
        {
            //ExStart
            //ExFor:CompositeNode.SelectNodes
            //ExSummary:Shows how to use an XPath expression to test whether a node is inside a field.
            Document doc = new Document(MyDir + "Mail merge destination - Northwind employees.docx");

            // The NodeList that results from this XPath expression will contain all nodes we find inside a field.
            // However, FieldStart and FieldEnd nodes can be on the list if there are nested fields in the path.
            // Currently does not find rare fields in which the FieldCode or FieldResult spans across multiple paragraphs.
            NodeList resultList =
                doc.SelectNodes("//FieldStart/following-sibling::node()[following-sibling::FieldEnd]");

            // Check if the specified run is one of the nodes that are inside the field.
            Console.WriteLine($"Contents of the first Run node that's part of a field: {resultList.First(n => n.NodeType == NodeType.Run).GetText().Trim()}");
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
            //ExSummary:Removes all smart tags from descendant nodes of a composite node.
            Document doc = new Document(MyDir + "Smart tags.doc");

            Assert.That(doc.GetChildNodes(NodeType.SmartTag, true).Count, Is.EqualTo(8));

            doc.RemoveSmartTags();

            Assert.That(doc.GetChildNodes(NodeType.SmartTag, true).Count, Is.EqualTo(0));
            //ExEnd
        }

        [Test]
        public void GetIndexOfNode()
        {
            //ExStart
            //ExFor:CompositeNode.IndexOf
            //ExSummary:Shows how to get the index of a given child node from its parent.
            Document doc = new Document(MyDir + "Rendering.docx");

            Body body = doc.FirstSection.Body;

            // Retrieve the index of the last paragraph in the body of the first section.
            Assert.That(body.GetChildNodes(NodeType.Any, false).IndexOf(body.LastParagraph), Is.EqualTo(24));
            //ExEnd
        }

        [Test]
        public void ConvertNodeToHtmlWithDefaultOptions()
        {
            //ExStart
            //ExFor:Node.ToString(SaveFormat)
            //ExFor:Node.ToString(SaveOptions)
            //ExSummary:Exports the content of a node to String in HTML format.
            Document doc = new Document(MyDir + "Document.docx");

            Node node = doc.LastSection.Body.LastParagraph;

            // When we call the ToString method using the html SaveFormat overload,
            // it converts the node's contents to their raw html representation.
            Assert.That(node.ToString(SaveFormat.Html), Is.EqualTo("<p style=\"margin-top:0pt; margin-bottom:8pt; line-height:108%; font-size:12pt\">" +
                            "<span style=\"font-family:'Times New Roman'\">Hello World!</span>" +
                            "</p>"));

            // We can also modify the result of this conversion using a SaveOptions object.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportRelativeFontSize = true;

            Assert.That(node.ToString(saveOptions), Is.EqualTo("<p style=\"margin-top:0pt; margin-bottom:8pt; line-height:108%\">" +
                            "<span style=\"font-family:'Times New Roman'\">Hello World!</span>" +
                            "</p>"));
            //ExEnd
        }

        [Test]
        public void TypedNodeCollectionToArray()
        {
            //ExStart
            //ExFor:ParagraphCollection.ToArray
            //ExSummary:Shows how to create an array from a NodeCollection.
            Document doc = new Document(MyDir + "Paragraphs.docx");

            Paragraph[] paras = doc.FirstSection.Body.Paragraphs.ToArray();

            Assert.That(paras.Length, Is.EqualTo(22));
            //ExEnd
        }

        [Test]
        public void NodeEnumerationHotRemove()
        {
            //ExStart
            //ExFor:ParagraphCollection.ToArray
            //ExSummary:Shows how to use "hot remove" to remove a node during enumeration.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("The first paragraph");
            builder.Writeln("The second paragraph");
            builder.Writeln("The third paragraph");
            builder.Writeln("The fourth paragraph");

            // Remove a node from the collection in the middle of an enumeration.
            foreach (Paragraph para in doc.FirstSection.Body.Paragraphs.ToArray())
                if (para.Range.Text.Contains("third"))
                    para.Remove();

            Assert.That(doc.GetText().Contains("The third paragraph"), Is.False);
            //ExEnd
        }

        //ExStart
        //ExFor:CompositeNode.CreateNavigator
        //ExSummary:Shows how to create an XPathNavigator, and then use it to traverse and read nodes.
        [Test] //ExSkip
        public void NodeXPathNavigator()
        {
            Document doc = new Document();
            XPathNavigator navigator = doc.CreateNavigator();

            if (navigator != null)
            {
                Assert.That(navigator.Name, Is.EqualTo("Document"));
                Assert.That(navigator.MoveToNext(), Is.EqualTo(false));
                Assert.That(navigator.SelectChildren(XPathNodeType.All).Count, Is.EqualTo(1));

                // The document tree has the document, first section,
                // body, and first paragraph as nodes, with each being an only child of the previous.
                // We can add a few more to give the tree some branches for the navigator to traverse.
                DocumentBuilder docBuilder = new DocumentBuilder(doc);
                docBuilder.Write("Section 1, Paragraph 1. ");
                docBuilder.InsertParagraph();
                docBuilder.Write("Section 1, Paragraph 2. ");
                doc.AppendChild(new Section(doc));
                docBuilder.MoveToSection(1);
                docBuilder.Write("Section 2, Paragraph 1. ");

                // Use our navigator to print a map of all the nodes in the document to the console.
                StringBuilder stringBuilder = new StringBuilder();
                MapDocument(navigator, stringBuilder, 0);
                Console.Write(stringBuilder.ToString());
                TestNodeXPathNavigator(stringBuilder.ToString(), doc); //ExSkip
            }
        }

        /// <summary>
        /// Traverses all children of a composite node and map the structure in the style of a directory tree.
        /// The amount of space indentation indicates depth relative to the initial node.
        /// Prints the text contents of the current node only if it is a Run.
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

        private void TestNodeXPathNavigator(string navigatorResult, Document doc)
        {
            foreach (Run run in doc.GetChildNodes(NodeType.Run, true).ToArray().OfType<Run>())
                Assert.That(navigatorResult.Contains(run.GetText().Trim()), Is.True);
        }

        //ExStart
        //ExFor:NodeChangingAction
        //ExFor:NodeChangingArgs.Action
        //ExFor:NodeChangingArgs.NewParent
        //ExFor:NodeChangingArgs.OldParent
        //ExSummary:Shows how to use a NodeChangingCallback to monitor changes to the document tree in real-time as we edit it.
        [Test] //ExSkip
        public void NodeChangingCallback()
        {
            Document doc = new Document();
            doc.NodeChangingCallback = new NodeChangingPrinter();

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndTable();

            builder.InsertImage(ImageDir + "Logo.jpg");

            builder.CurrentParagraph.ParentNode.RemoveAllChildren();
        }

        /// <summary>
        /// Prints every node insertion/removal as it takes place in the document.
        /// </summary>
        private class NodeChangingPrinter : INodeChangingCallback
        {
            void INodeChangingCallback.NodeInserting(NodeChangingArgs args)
            {
                Assert.That(args.Action, Is.EqualTo(NodeChangingAction.Insert));
                Assert.That(args.OldParent, Is.EqualTo(null));
            }

            void INodeChangingCallback.NodeInserted(NodeChangingArgs args)
            {
                Assert.That(args.Action, Is.EqualTo(NodeChangingAction.Insert));
                Assert.That(args.NewParent, Is.Not.Null);

                Console.WriteLine("Inserted node:");
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
                Assert.That(args.Action, Is.EqualTo(NodeChangingAction.Remove));
            }

            void INodeChangingCallback.NodeRemoved(NodeChangingArgs args)
            {
                Assert.That(args.Action, Is.EqualTo(NodeChangingAction.Remove));
                Assert.That(args.NewParent, Is.Null);

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

            // Add text to the document by inserting Runs using a DocumentBuilder.
            builder.Write("Run 1. ");
            builder.Write("Run 2. ");

            // Every invocation of the "Write" method creates a new Run,
            // which then appears in the parent Paragraph's RunCollection.
            RunCollection runs = doc.FirstSection.Body.FirstParagraph.Runs;

            Assert.That(runs.Count, Is.EqualTo(2));

            // We can also insert a node into the RunCollection manually.
            Run newRun = new Run(doc, "Run 3. ");
            runs.Insert(3, newRun);

            Assert.That(runs.Contains(newRun), Is.True);
            Assert.That(doc.GetText().Trim(), Is.EqualTo("Run 1. Run 2. Run 3."));

            // Access individual runs and remove them to remove their text from the document.
            Run run = runs[1];
            runs.Remove(run);

            Assert.That(doc.GetText().Trim(), Is.EqualTo("Run 1. Run 3."));
            Assert.That(run, Is.Not.Null);
            Assert.That(runs.Contains(run), Is.False);
            //ExEnd
        }

        [Test]
        public void NodeList()
        {
            //ExStart
            //ExFor:NodeList.Count
            //ExFor:NodeList.Item(Int32)
            //ExSummary:Shows how to use XPaths to navigate a NodeList.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert some nodes with a DocumentBuilder.
            builder.Writeln("Hello world!");

            builder.StartTable();
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndTable();

            builder.InsertImage(ImageDir + "Logo.jpg");

            // Our document contains three Run nodes.
            NodeList nodeList = doc.SelectNodes("//Run");

            Assert.That(nodeList.Count, Is.EqualTo(3));
            Assert.That(nodeList.Any(n => n.GetText().Trim() == "Hello world!"), Is.True);
            Assert.That(nodeList.Any(n => n.GetText().Trim() == "Cell 1"), Is.True);
            Assert.That(nodeList.Any(n => n.GetText().Trim() == "Cell 2"), Is.True);

            // Use a double forward slash to select all Run nodes
            // that are indirect descendants of a Table node, which would be the runs inside the two cells we inserted.
            nodeList = doc.SelectNodes("//Table//Run");

            Assert.That(nodeList.Count, Is.EqualTo(2));
            Assert.That(nodeList.Any(n => n.GetText().Trim() == "Cell 1"), Is.True);
            Assert.That(nodeList.Any(n => n.GetText().Trim() == "Cell 2"), Is.True);

            // Single forward slashes specify direct descendant relationships,
            // which we skipped when we used double slashes.
            Assert.That(doc.SelectNodes("//Table/Row/Cell/Paragraph/Run"), Is.EqualTo(doc.SelectNodes("//Table//Run")));

            // Access the shape that contains the image we inserted.
            nodeList = doc.SelectNodes("//Shape");

            Assert.That(nodeList.Count, Is.EqualTo(1));

            Shape shape = (Shape)nodeList[0];
            Assert.That(shape.HasImage, Is.True);
            //ExEnd
        }
    }
}
