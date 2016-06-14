using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;
using AWords = Aspose.Words;


namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Node
{
    class ExNode
    {
        public static void Run()
        {
            // The following method shows how to use the NodeType enumeration.
            UseNodeType();
            // The following method shows how to access the parent node.
            GetParentNode();
            // The following method shows that when you create any node, it requires a document that will own the node.
            OwnerDocument();
            // Shows how to extract a specific child node from a CompositeNode by using the GetChild method and passing the NodeType and index.
            EnumerateChildNodes();
            // Shows how to enumerate immediate children of a CompositeNode using indexed access.
            IndexChildNodes();
            // Shows how to efficiently visit all direct and indirect children of a composite node.
            RecurseAllNodes();
            // Demonstrates how to use typed properties to access nodes of the document tree.
            TypedAccess();
            // The following method shows how to creates and adds a paragraph node.
            CreateAndAddParagraphNode();
        }
        public static void UseNodeType()
        {
            //ExStart:UseNodeType            
            Document doc = new Document();
            // Returns NodeType.Document
            NodeType type = doc.NodeType;
            //ExEnd:UseNodeType
        }
        public static void GetParentNode()
        {
            //ExStart:GetParentNode           
            // Create a new empty document. It has one section.
            Document doc = new Document();
            // The section is the first child node of the document.
            Node section = doc.FirstChild;
            // The section's parent node is the document.
            Console.WriteLine("Section parent is the document: " + (doc == section.ParentNode));
            //ExEnd:GetParentNode           
        }
        public static void OwnerDocument()
        {
            //ExStart:OwnerDocument            
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
            //ExEnd:OwnerDocument
        }
        public static void EnumerateChildNodes()
        {
            //ExStart:EnumerateChildNodes 
            Document doc = new Document();          
            Paragraph paragraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
            
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
            //ExEnd:EnumerateChildNodes
        }
        public static void IndexChildNodes()
        {
            //ExStart:IndexChildNodes
            Document doc = new Document();
            Paragraph paragraph = (Paragraph)doc.GetChild(NodeType.Paragraph, 0, true);
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
            //ExEnd:IndexChildNodes
        }
        //ExStart:RecurseAllNodes
        public static void RecurseAllNodes()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithNode();
            // Open a document.
            Document doc = new Document(dataDir + "Node.RecurseAllNodes.doc");

            // Invoke the recursive function that will walk the tree.
            TraverseAllNodes(doc);
        }

        /// <summary>
        /// A simple function that will walk through all children of a specified node recursively 
        /// and print the type of each node to the screen.
        /// </summary>
        public static void TraverseAllNodes(CompositeNode parentNode)
        {
            // This is the most efficient way to loop through immediate children of a node.
            for (Node childNode = parentNode.FirstChild; childNode != null; childNode = childNode.NextSibling)
            {
                // Do some useful work.
                Console.WriteLine(Node.NodeTypeToString(childNode.NodeType));

                // Recurse into the node if it is a composite node.
                if (childNode.IsComposite)
                    TraverseAllNodes((CompositeNode)childNode);
            }
        }
        //ExEnd:RecurseAllNodes
        public static void TypedAccess()
        {
            //ExStart:TypedAccess
            Document doc = new Document();            
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
            //ExEnd:TypedAccess
        }
        public static void CreateAndAddParagraphNode()
        {
            //ExStart:CreateAndAddParagraphNode
            Document doc = new Document();
            Paragraph para = new Paragraph(doc);
            AWords.Section section = doc.LastSection;
            section.Body.AppendChild(para);
            //ExEnd:CreateAndAddParagraphNode
        }
    }

}
