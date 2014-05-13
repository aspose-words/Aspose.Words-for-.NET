//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Text;

using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Tables;

namespace Aspose.Words.Layout
{
    /// <summary>
    /// Provides an API wrapper for the LayoutEnumerator class to access the page layout entities of a document presented in
    /// a object model like design.
    /// </summary>
    public class RenderedDocument : LayoutEntity
    {
        /// <summary>
        /// Creates a new instance from the supplied Aspose.Words.Document class.
        /// </summary>
        /// <param name="document">A document whose page layout model to enumerate.</param>
        /// <remarks><para>If page layout model of the document hasn't been built the enumerator calls <see cref="Document.UpdatePageLayout"/> to build it.</para>
        /// <para>Whenever document is updated and new page layout model is created, a new RenderedDocument instance must be used to access the changes.</para></remarks>
        public RenderedDocument(Document doc)
        {
            mLayoutCollector = new LayoutCollector(doc);
            mEnumerator = new LayoutEnumerator(doc);
            ProcessLayoutElements(this);
            LinkLayoutMarkersToNodes(doc);
            CollectLinesAndAddToMarkers();
        }

        /// <summary>
        /// Provides access to the pages of a document.
        /// </summary>
        public LayoutCollection<RenderedPage> Pages
        {
            get
            {
                return GetChildNodes<RenderedPage>();
            }
        }

        /// <summary>
        /// Returns all the layout entities of the specified node.
        /// </summary>
        /// <remarks>Note that this method does not work with Run nodes or nodes in the header or footer.</remarks>
        public LayoutCollection<LayoutEntity> GetLayoutEntitiesOfNode(Node node)
        {
            if (!mLayoutCollector.Document.Equals(node.Document))
                throw new ArgumentException("Node does not belong to the same document which was rendered.");

            if (node.NodeType == NodeType.Document)
                return new LayoutCollection<LayoutEntity>(mChildEntities);
           
            List<LayoutEntity> entities = new List<LayoutEntity>();

            // Retrieve all entities from the layout document (inversion of LayoutEntityType.None).
            foreach (LayoutEntity entity in GetChildEntities(~LayoutEntityType.None, true))
            {
                if (entity.ParentNode == node)
                    entities.Add(entity);

                // There is no table entity in rendered output so manually check if rows belong to a table node.
                if (entity.Type == LayoutEntityType.Row)
                {
                    RenderedRow row = (RenderedRow)entity;
                    if(row.Table == node)
                        entities.Add(entity);
                }
            }

            return new LayoutCollection<LayoutEntity>(entities);
        }

        private void ProcessLayoutElements(LayoutEntity current)
        {
            do
            {
                LayoutEntity child = current.AddChildEntity(mEnumerator);

                if (mEnumerator.MoveFirstChild())
                {
                    current = child;

                    ProcessLayoutElements(current);
                    mEnumerator.MoveParent();

                    current = current.Parent;
                }
            } while (mEnumerator.MoveNext());
        }

        private void CollectLinesAndAddToMarkers()
        {
            CollectLinesOfMarkersCore(LayoutEntityType.Column);
            CollectLinesOfMarkersCore(LayoutEntityType.Comment);
        }

        private void CollectLinesOfMarkersCore(LayoutEntityType type)
        {
            List<RenderedLine> collectedLines = new List<RenderedLine>();

            foreach (RenderedPage page in Pages)
            {
                foreach (LayoutEntity story in page.GetChildEntities(type, false))
                {
                    foreach (RenderedLine line in story.GetChildEntities(LayoutEntityType.Line, true))
                    {
                        collectedLines.Add(line);
                        foreach (RenderedSpan span in line.Spans)
                        {
                            if (mLayoutToNodeLookup.ContainsKey(span.LayoutObject))
                            {
                                if (span.Kind == "PARAGRAPH" || span.Kind == "ROW" || span.Kind == "CELL" || span.Kind == "SECTION")
                                {
                                    Node node = mLayoutToNodeLookup[span.LayoutObject];

                                    if (node.NodeType == NodeType.Row)
                                        node = ((Row)node).LastCell.LastParagraph;

                                    foreach (RenderedLine collectedLine in collectedLines)
                                        collectedLine.SetParentNode(node);

                                    collectedLines = new List<RenderedLine>();
                                }
                                else
                                {
                                    span.SetParentNode(mLayoutToNodeLookup[span.LayoutObject]);
                                }
                            }
                        }
                    }
                }
            }
        }

        private void LinkLayoutMarkersToNodes(Document doc)
        {
            foreach (Node node in doc.GetChildNodes(NodeType.Any, true))
            {
                object entity = mLayoutCollector.GetEntity(node);

                if (entity != null)
                    mLayoutToNodeLookup.Add(entity, node);
            }
        }

        private LayoutCollector mLayoutCollector;
        private LayoutEnumerator mEnumerator;
        private static Dictionary<object, Node> mLayoutToNodeLookup =
                new Dictionary<object, Node>();
    }
}