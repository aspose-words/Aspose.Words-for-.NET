// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;

using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExNode : ApiExampleBase
    {
        [Test]
        public void UseNodeType()
        {
            //ExStart
            //ExFor:NodeType
            //ExId:UseNodeType
            //ExSummary:The following example shows how to use the NodeType enumeration.
            Document doc = new Document();

            // Returns NodeType.Document
            NodeType type = doc.NodeType;
            //ExEnd
        }

        [Test]
        public void CloneCompositeNode()
        {
            //ExStart
            //ExFor:Node
            //ExFor:Node.Clone
            //ExSummary:Shows how to clone composite nodes with and without their child nodes.
            // Create a new empty document.
            Document doc = new Document();

            // Add some text to the first paragraph
            Paragraph para = doc.FirstSection.Body.FirstParagraph;
            para.AppendChild(new Run(doc, "Some text"));

            // Clone the paragraph and the child nodes.
            Node cloneWithChildren = para.Clone(true);
            // Only clone the paragraph and no child nodes.
            Node cloneWithoutChildren = para.Clone(false);
            //ExEnd

            Assert.IsTrue(((CompositeNode)cloneWithChildren).HasChildNodes);
            Assert.IsFalse(((CompositeNode)cloneWithoutChildren).HasChildNodes);
        }
        
        [Test]
	    public void GetParentNode()
	    {
            //ExStart
            //ExFor:Node.ParentNode
            //ExId:AccessParentNode
            //ExSummary:Shows how to access the parent node.
            // Create a new empty document. It has one section.
            Document doc = new Document();

            // The section is the first child node of the document.
            Node section = doc.FirstChild;

            // The section's parent node is the document.
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
            //ExId:CreatingNodeRequiresDocument
            //ExSummary:Shows that when you create any node, it requires a document that will own the node.
            // Open a file from disk.
            Document doc = new Document();

            // Creating a new node of any type requires a document passed into the constructor.
            Paragraph para = new Paragraph(doc);

            // The new paragraph node does not yet have a parent.
            Console.WriteLine("Paragraph has no parent node: " + (para.ParentNode == null));
           
            // But the paragraph node knows its document.
            Console.WriteLine("Both nodes' documents are the same: " + (para.Document == doc));

            // The fact that a node always belongs to a document allows us to access and modify 
            // properties that reference the document-wide data such as styles or lists.
            para.ParagraphFormat.StyleName = "Heading 1";

            // Now add the paragraph to the main text of the first section.
            doc.FirstSection.Body.AppendChild(para);

            // The paragraph node is now a child of the Body node.
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
            //ExFor:CompositeNode
            //ExFor:CompositeNode.GetChild
            //ExSummary:Shows how to extract a specific child node from a CompositeNode by using the GetChild method and passing the NodeType and index.
            Paragraph paragraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
            //ExEnd

            //ExStart
            //ExFor:CompositeNode.ChildNodes
            //ExFor:CompositeNode.GetEnumerator
            //ExId:ChildNodesForEach
            //ExSummary:Shows how to enumerate immediate children of a CompositeNode using the enumerator provided by the ChildNodes collection.
            NodeCollection children = paragraph.ChildNodes;
            foreach (Node child in children)
            {
                // Paragraph may contain children of various types such as runs, shapes and so on.
                if (child.NodeType.Equals(NodeType.Run))
                {
                    // Say we found the node that we want, do something useful.
                    Run run = (Run)child;
                    Console.WriteLine(run.Text);
                }
            }
            //ExEnd
        }

        [Test]
        public void IndexChildNodes()
        {
            Document doc = new Document();
            Paragraph paragraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);

            //ExStart
            //ExFor:NodeCollection.Count
            //ExFor:NodeCollection.Item
            //ExId:ChildNodesIndexer
            //ExSummary:Shows how to enumerate immediate children of a CompositeNode using indexed access.
            NodeCollection children = paragraph.ChildNodes;
            for (int i = 0; i < children.Count; i++)
            {
                Node child = children[i];

                // Paragraph may contain children of various types such as runs, shapes and so on.
                if (child.NodeType.Equals(NodeType.Run))
                {
                    // Say we found the node that we want, do something useful.
                    Run run = (Run)child;
                    Console.WriteLine(run.Text);
                }
            }
            //ExEnd
        }

        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void RecurseAllNodesCaller()
        {
            this.RecurseAllNodes();
        }
        
        //ExStart
        //ExFor:Node.NextSibling
        //ExFor:CompositeNode.FirstChild
        //ExFor:Node.IsComposite
        //ExFor:CompositeNode.IsComposite
        //ExFor:Node.NodeTypeToString
        //ExId:RecurseAllNodes            
        //ExSummary:Shows how to efficiently visit all direct and indirect children of a composite node.
        public void RecurseAllNodes()
        {
            // Open a document.
            Document doc = new Document(MyDir + "Node.RecurseAllNodes.doc");

            // Invoke the recursive function that will walk the tree.
            this.TraverseAllNodes(doc);
        }

        /// <summary>
        /// A simple function that will walk through all children of a specified node recursively 
        /// and print the type of each node to the screen.
        /// </summary>
        public void TraverseAllNodes(CompositeNode parentNode)
        {
            // This is the most efficient way to loop through immediate children of a node.
            for (Node childNode = parentNode.FirstChild; childNode != null; childNode = childNode.NextSibling)
            {
                // Do some useful work.
                Console.WriteLine(Node.NodeTypeToString(childNode.NodeType));

                // Recurse into the node if it is a composite node.
                if (childNode.IsComposite)
                    this.TraverseAllNodes((CompositeNode)childNode);
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
            //ExSummary:Shows how to remove all nodes of a specific type from a composite node. In this example we remove tables from a section body.
            // Get the section that we want to work on.
            Section section = doc.Sections[0];
            Body body = section.Body;

            // Select the first child node in the body.
            Node curNode = body.FirstChild;

            while (curNode != null)
            {
                // Save the pointer to the next sibling node because if the current 
                // node is removed from the parent in the next step, we will have 
                // no way of finding the next node to continue the loop.
                Node nextNode = curNode.NextSibling;

                // A section body can contain Paragraph and Table nodes.
                // If the node is a Table, remove it from the parent.
                if (curNode.NodeType.Equals(NodeType.Table))
                    curNode.Remove();

                // Continue going through child nodes until null (no more siblings) is reached.
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
            //ExSummary:Shows how to enumerate immediate child nodes of a composite node using NextSibling. In this example we enumerate all paragraphs of a section body.
            // Get the section that we want to work on.
            Section section = doc.Sections[0];
            Body body = section.Body;

            // Loop starting from the first child until we reach null.
            for (Node node = body.FirstChild; node != null; node = node.NextSibling)
            {
                // Output the types of the nodes that we come across.
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
            //ExId:TypedPropertiesAccess
            //ExSummary:Demonstrates how to use typed properties to access nodes of the document tree.
            // Quick typed access to the first child Section node of the Document.
            Section section = doc.FirstSection;

            // Quick typed access to the Body child node of the Section.
            Body body = section.Body;

            // Quick typed access to all Table child nodes contained in the Body.
            TableCollection tables = body.Tables;

            foreach (Table table in tables)
            {
                // Quick typed access to the first row of the table.
                if (table.FirstRow != null)
                    table.FirstRow.Remove();

                // Quick typed access to the last row of the table.
                if (table.LastRow != null)
                    table.LastRow.Remove();
            }
            //ExEnd
        }

        [Test]
        public void UpdateFieldsInRange()
        {
            Document doc = new Document();

            //ExStart
            //ExFor:Range.UpdateFields
            //ExSummary:Demonstrates how to update document fields in the body of the first section only.
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
            //ExSummary:Demonstrates use of methods of Node and CompositeNode to remove a section before the last section in the document.
            // Document is a CompositeNode and LastChild returns the last child node in the Document node.
            // Since the Document can contain only Section nodes, the last child is the last section.
            Node lastSection = doc.LastChild;
            
            // Each node knows its next and previous sibling nodes.
            // Previous sibling of a section is a section before the specified section.
            // If the node is the first child, PreviousSibling will return null.
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
            //ExSummary:Shows how to select certain nodes by using an XPath expression.
            Document doc = new Document(MyDir + "Table.Document.doc");

            // This expression will extract all paragraph nodes which are descendants of any table node in the document.
            // This will return any paragraphs which are in a table.
            NodeList nodeList = doc.SelectNodes("//Table//Paragraph");

            // This expression will select any paragraphs that are direct children of any body node in the document.
            nodeList = doc.SelectNodes("//Body/Paragraph");

            // Use SelectSingleNode to select the first result of the same expression as above.
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
            // Let's pick a document we know has some fields in.
            Document doc = new Document(MyDir + "MailMerge.MergeImage.doc");

            // Let's say we want to check if the Run below is inside a field.
            Run run = (Run)doc.GetChild(NodeType.Run, 5, true);

            // Evaluate the XPath expression. The resulting NodeList will contain all nodes found inside a field a field (between FieldStart 
            // and FieldEnd exclusive). There can however be FieldStart and FieldEnd nodes in the list if there are nested fields 
            // in the path. Currently does not find rare fields in which the FieldCode or FieldResult spans across multiple paragraphs.
            NodeList resultList = doc.SelectNodes("//FieldStart/following-sibling::node()[following-sibling::FieldEnd]");

            // Check if the specified run is one of the nodes that are inside the field.
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
            //ExStart
            //ExId:CreateAndAddParagraphNode
            //ExSummary:Creates and adds a paragraph node.
            Document doc = new Document();

            Paragraph para = new Paragraph(doc);

            Section section = doc.LastSection;
            section.Body.AppendChild(para);
            //ExEnd
        }

        [Test]
        public void RemoveSmartTagsFromCompositeNode()
        {
            //ExStart
            //ExFor:CompositeNode.RemoveSmartTags
            //ExSummary:Removes all smart tags from descendant nodes of the composite node.
            Document doc = new Document(MyDir + "Document.doc");

            // Remove smart tags from the first paragraph in the document.
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

            // Get the body of the first section in the document.
            Body body = doc.FirstSection.Body;
            // Retrieve the index of the last paragraph in the body.
            int index = body.ChildNodes.IndexOf(body.LastParagraph);
            //ExEnd

            // Verify that the index is correct.
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
            //ExId:GetNodeTypeEnums
            //ExSummary:Shows how to retrieve the NodeType enumeration of nodes.
            Document doc = new Document(MyDir + "Document.doc");

            // Let's pick a node that we can't be quite sure of what type it is.
            // In this case lets pick the first node of the first paragraph in the body of the document
            Node node = doc.FirstSection.Body.FirstParagraph.FirstChild;
            Console.WriteLine("NodeType of first child: " + Node.NodeTypeToString(node.NodeType));

            // This time let's pick a node that we know the type of. Create a new paragraph and a table node.
            Paragraph para = new Paragraph(doc);
            Table table = new Table(doc);

            // Access to NodeType for typed nodes will always return their specific NodeType. 
            // i.e A paragraph node will always return NodeType.Paragraph, a table node will always return NodeType.Table.
            Console.WriteLine("NodeType of Paragraph: " + Node.NodeTypeToString(para.NodeType));
            Console.WriteLine("NodeType of Table: " + Node.NodeTypeToString(table.NodeType));
            //ExEnd
        }

        [Test]
        public void ConvertNodeToHtmlWithDefaultOptions()
        {
            //ExStart
            //ExFor:Node.ToString(SaveFormat)
            //ExSummary:Exports the content of a node to string in HTML format using default options.
            Document doc = new Document(MyDir + "Document.doc");

            // Extract the last paragraph in the document to convert to HTML.
            Node node = doc.LastSection.Body.LastParagraph;

            // When ToString is called using the SaveFormat overload then conversion is executed using default save options. 
            // When saving to HTML using default options the following settings are set:
            //   ExportImagesAsBase64 = true
            //   CssStyleSheetType = CssStyleSheetType.Inline
            //   ExportFontResources = false
            string nodeAsHtml = node.ToString(SaveFormat.Html);
            //ExEnd

            Assert.AreEqual("<p style=\"margin-top:0pt; margin-bottom:0pt; font-size:12pt\"><span style=\"font-family:'Times New Roman'\">Hello World!</span></p>", nodeAsHtml);
        }

        [Test]
        public void ConvertNodeToHtmlWithSaveOptions()
        {
            //ExStart
            //ExFor:Node.ToString(SaveOptions)
            //ExSummary:Exports the content of a node to string in HTML format using custom specified options.
            Document doc = new Document(MyDir + "Document.doc");

            // Extract the last paragraph in the document to convert to HTML.
            Node node = doc.LastSection.Body.LastParagraph;

            // Create an instance of HtmlSaveOptions and set a few options.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportHeadersFootersMode = ExportHeadersFootersMode.PerSection;
            saveOptions.ExportRelativeFontSize = true;

            // Convert the document to HTML and return as a string. Pass the instance of HtmlSaveOptions to
            // to use the specified options during the conversion.
            string nodeAsHtml = node.ToString(saveOptions);
            //ExEnd

            Assert.AreEqual("<p style=\"margin-top:0pt; margin-bottom:0pt\"><span style=\"font-family:'Times New Roman'\">Hello World!</span></p>", nodeAsHtml);
        }

        [Test]
        public void TypedNodeCollectionToArray()
        {
            Document doc = new Document();

            //ExStart
            //ExFor:ParagraphCollection.ToArray
            //ExSummary:Demonstrates typed implementations of ToArray on classes derived from NodeCollection.
            // You can use ToArray to return a typed array of nodes.
            Paragraph[] paras = doc.FirstSection.Body.Paragraphs.ToArray();
            //ExEnd

            Assert.Greater(paras.Length,  0);
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

            // Hot remove allows a node to be removed from a live collection and have the enumeration continue.
            foreach (Paragraph para in builder.Document.FirstSection.Body.GetChildNodes(NodeType.Paragraph, true))
            {
                if (para.Range.Text.Contains("third"))
                {
                    // Enumeration will continue even after this node is removed.
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

            // This causes unexpected behavior, the fourth pargraph in the collection is not visited.
            foreach (Paragraph para in builder.Document.FirstSection.Body.GetChildNodes(NodeType.Paragraph, true))
            {
                if (para.Range.Text.Contains("third"))
                {
                    para.PreviousSibling.Remove();
                    para.Remove();
                }
            }
            //ExEnd
        }
    }
}
