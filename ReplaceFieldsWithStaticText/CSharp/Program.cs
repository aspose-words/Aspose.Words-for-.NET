//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using System.Reflection;
using System.Collections;

using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;
using System.Diagnostics;

namespace ReplaceFieldsWithStaticText
{
    class Program
    {
        public static void Main(string[] args)
        {
            // Sample infrastructure.
            string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar;
            string dataDir = new Uri(new Uri(exeDir), @"../../Data/").LocalPath;

            // Call a method to show how to convert all IF fields in a document to static text.
            ConvertFieldsInDocument(dataDir);
            // Reload the document and this time convert all PAGE fields only encountered in the first body of the document.
            ConvertFieldsInBody(dataDir);
            // Reload the document again and convert only the IF field in the last paragraph to static text.
            ConvertFieldsInParagraph(dataDir);
        }

        //ExStart:
        //ExFor:DocumentVisitor.VisitTableStart(Aspose.Words.Tables.Table)
        //ExId:ConvertFieldsToStaticText
        //ExSummary:This class provides a static method convert fields of a particular type to static text.
        public class FieldsHelper : DocumentVisitor
        {
            /// <summary>
            /// Converts any fields of the specified type found in the descendants of the node into static text.
            /// </summary>
            /// <param name="compositeNode">The node in which all descendants of the specified FieldType will be converted to static text.</param>
            /// <param name="targetFieldType">The FieldType of the field to convert to static text.</param>
            public static void ConvertFieldsToStaticText(CompositeNode compositeNode, FieldType targetFieldType)
            {
                string originalNodeText = compositeNode.ToTxt(); //ExSkip
                FieldsHelper helper = new FieldsHelper(targetFieldType);
                compositeNode.Accept(helper);

                Debug.Assert(originalNodeText.Equals(compositeNode.ToTxt()), "Error: Text of the node converted differs from the original"); //ExSkip
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
        //ExEnd

        public static void ConvertFieldsInDocument(string dataDir)
        {
            //ExStart:
            //ExId:FieldsToStaticTextDocument
            //ExSummary:Shows how to convert all fields of a specified type in a document to static text.
            Document doc = new Document(dataDir + "TestFile.doc");

            // Pass the appropriate parameters to convert all IF fields encountered in the document (including headers and footers) to static text.
            FieldsHelper.ConvertFieldsToStaticText(doc, FieldType.FieldIf);

            // Save the document with fields transformed to disk.
            doc.Save(dataDir + "TestFileDocument Out.doc");
            //ExEnd
        }

        public static void ConvertFieldsInBody(string dataDir)
        {
            //ExStart:
            //ExId:FieldsToStaticTextBody
            //ExSummary:Shows how to convert all fields of a specified type in a body of a document to static text.
            Document doc = new Document(dataDir + "TestFile.doc");

            // Pass the appropriate parameters to convert PAGE fields encountered to static text only in the body of the first section.
            FieldsHelper.ConvertFieldsToStaticText(doc.FirstSection.Body, FieldType.FieldPage);

            // Save the document with fields transformed to disk.
            doc.Save(dataDir + "TestFileBody Out.doc");
            //ExEnd
        }

        public static void ConvertFieldsInParagraph(string dataDir)
        {
            //ExStart:
            //ExId:FieldsToStaticTextParagraph
            //ExSummary:Shows how to convert all fields of a specified type in a paragraph to static text.
            Document doc = new Document(dataDir + "TestFile.doc");

            // Pass the appropriate parameters to convert all IF fields to static text that are encountered only in the last 
            // paragraph of the document.
            FieldsHelper.ConvertFieldsToStaticText(doc.FirstSection.Body.LastParagraph, FieldType.FieldIf);

            // Save the document with fields transformed to disk.
            doc.Save(dataDir + "TestFileParagraph Out.doc");
            //ExEnd
        }

    }
}
