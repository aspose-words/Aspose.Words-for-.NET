using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;
using System.Diagnostics;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    //ExStart:FieldsHelper
    public class FieldsHelper : DocumentVisitor
    {
        /// <summary>
        /// Converts any fields of the specified type found in the descendants of the node into static text.
        /// </summary>
        /// <param name="compositeNode">The node in which all descendants of the specified FieldType will be converted to static text.</param>
        /// <param name="targetFieldType">The FieldType of the field to convert to static text.</param>
        public static void ConvertFieldsToStaticText(CompositeNode compositeNode, FieldType targetFieldType)
        {
            string originalNodeText = compositeNode.ToString(SaveFormat.Text); //ExSkip
            FieldsHelper helper = new FieldsHelper(targetFieldType);
            compositeNode.Accept(helper);

            Debug.Assert(originalNodeText.Equals(compositeNode.ToString(SaveFormat.Text)), "Error: Text of the node converted differs from the original"); //ExSkip
            foreach (Node node in compositeNode.GetChildNodes(NodeType.Any, true)) //ExSkip
                Debug.Assert(!(node is FieldChar && ((FieldChar)node).FieldType.Equals(targetFieldType)), "Error: A field node that should be removed still remains."); //ExSkip         
        }

        private FieldsHelper(FieldType targetFieldType)
        {
            mTargetFieldType = targetFieldType;
        }

        public override VisitorAction VisitFieldStart(FieldStart fieldStart)
        {
            // We must keep track of the starts and ends of fields incase of any nested fields.
            if (fieldStart.FieldType.Equals(mTargetFieldType))
            {
                mFieldDepth++;
                fieldStart.Remove();
            }
            else
            {
                // This removes the field start if it's inside a field that is being converted.
                CheckDepthAndRemoveNode(fieldStart);
            }

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitFieldSeparator(FieldSeparator fieldSeparator)
        {
            // When visiting a field separator we should decrease the depth level.
            if (fieldSeparator.FieldType.Equals(mTargetFieldType))
            {
                mFieldDepth--;
                fieldSeparator.Remove();
            }
            else
            {
                // This removes the field separator if it's inside a field that is being converted.
                CheckDepthAndRemoveNode(fieldSeparator);
            }

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitFieldEnd(FieldEnd fieldEnd)
        {
            if (fieldEnd.FieldType.Equals(mTargetFieldType))
                fieldEnd.Remove();
            else
                CheckDepthAndRemoveNode(fieldEnd); // This removes the field end if it's inside a field that is being converted.

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitRun(Run run)
        {
            // Remove the run if it is between the FieldStart and FieldSeparator of the field being converted.
            CheckDepthAndRemoveNode(run);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitParagraphEnd(Paragraph paragraph)
        {
            if (mFieldDepth > 0)
            {
                // The field code that is being converted continues onto another paragraph. We 
                // need to copy the remaining content from this paragraph onto the next paragraph.
                Node nextParagraph = paragraph.NextSibling;

                // Skip ahead to the next available paragraph.
                while (nextParagraph != null && nextParagraph.NodeType != NodeType.Paragraph)
                    nextParagraph = nextParagraph.NextSibling;

                // Copy all of the nodes over. Keep a list of these nodes so we know not to remove them.
                while (paragraph.HasChildNodes)
                {
                    mNodesToSkip.Add(paragraph.LastChild);
                    ((Paragraph)nextParagraph).PrependChild(paragraph.LastChild);
                }

                paragraph.Remove();
            }

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitTableStart(Table table)
        {
            CheckDepthAndRemoveNode(table);

            return VisitorAction.Continue;
        }

        /// <summary>
        /// Checks whether the node is inside a field or should be skipped and then removes it if necessary.
        /// </summary>
        private void CheckDepthAndRemoveNode(Node node)
        {
            if (mFieldDepth > 0 && !mNodesToSkip.Contains(node))
                node.Remove();
        }

        private int mFieldDepth = 0;
        private ArrayList mNodesToSkip = new ArrayList();
        private FieldType mTargetFieldType;
    }
    //ExEnd:FieldsHelper
}
